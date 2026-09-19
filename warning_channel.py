"""The warnings channel a builder grows: what it worked around, as data.

The PowerPoint tool has reported its own workarounds since #122, and #114
gave Word and Excel the same channel. This module is the part all of them
share: the severity vocabulary, one warning record, and the collector that a
single build writes into.

A warning is not an error. The file was produced; something in it is not
exactly what the caller asked for, and the caller — usually a model about to
write its next call — is the only one who can fix it. Left in the server log,
that information never reaches them.

Each :class:`DocumentWarning` carries:

* ``code`` — a stable identifier, the thing to branch on;
* ``severity`` — how much it cost, derived from the code so one code always
  means one severity;
* ``message`` — one sentence, naming what to change;
* ``location`` — where it happened, as the tool counts positions: a source
  ``line`` for Word, a ``sheet`` and ``cell`` for Excel. It is spread into
  ``as_dict()`` rather than nested, so a caller reads ``warning["line"]``.

Severity is about the document, not about how unusual the situation is:

``error``
    Something the caller asked for is not in the file. A dropped block, an
    image that would not load, a formula stored as text.
``warning``
    It is in the file, but not as asked: a renamed header, a substituted
    style, a styling directive that was skipped.
``info``
    A substitution the caller probably does not mind.

One :class:`WarningChannel` belongs to one build. Builds run concurrently on
worker threads (see ``async_runner``), so the channel is passed down the call
chain as an argument — never a module-level list.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterator, List, Mapping, Tuple

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"
SEVERITY_INFO = "info"

#: Every severity a code may map to.
SEVERITIES = (SEVERITY_ERROR, SEVERITY_WARNING, SEVERITY_INFO)

#: The one code the channel raises about itself, when a build produced more
#: warnings than it is willing to carry back.
WARNINGS_TRUNCATED = "warnings_truncated"

#: How many distinct warnings one build reports. A markdown file of prose fed
#: to the Excel tool drops a line per line; a response carrying thousands of
#: them is no more useful than one carrying fifty, and costs the caller its
#: context window.
DEFAULT_LIMIT = 50


@dataclass(frozen=True)
class DocumentWarning:
    """One thing a builder worked around, and where."""

    code: str
    message: str
    severity: str = SEVERITY_WARNING
    location: Mapping[str, Any] = field(default_factory=dict)

    def __str__(self) -> str:
        """The log line: ``line 42: …``, ``sheet 'Data', cell C7: …``."""
        where = ", ".join(f"{key} {value}" for key, value in self.location.items())
        return f"{where}: {self.message}" if where else self.message

    def as_dict(self) -> Dict[str, Any]:
        """The shape the tool returns; the location keys are spread in."""
        return {
            "code": self.code,
            "severity": self.severity,
            "message": self.message,
            **dict(self.location),
        }


class WarningChannel:
    """The warnings of one build, in the order they happened.

    *severities* maps a code to its severity — one table per tool, so a code
    cannot mean an error in one place and an info in another. An unregistered
    code is a programming error, but not one worth failing a finished document
    over: it is reported at ``warning`` severity, and each tool's
    ``test_every_code_has_a_severity`` keeps its table honest.

    Two properties make the channel safe to call from inside a loop over every
    cell or every line:

    * the same ``(code, message, location)`` is recorded once — a template
      missing ``List Number`` warns once, not once per list item;
    * after *limit* distinct warnings the rest are counted, not kept, and a
      final :data:`WARNINGS_TRUNCATED` entry says how many were dropped.
    """

    def __init__(self, severities: Mapping[str, str], limit: int = DEFAULT_LIMIT):
        self._severities = severities
        self._limit = limit
        self._warnings: List[DocumentWarning] = []
        self._seen: set[Tuple[Any, ...]] = set()
        self._suppressed = 0

    def add(self, code: str, message: str, **location: Any) -> None:
        """Record *message* under *code*, at *location* (``line=``, ``cell=``…).

        Location entries whose value is ``None`` are dropped, so a call site
        that does not know the line can pass ``line=None`` unconditionally.
        """
        where = {key: value for key, value in location.items() if value is not None}
        key = (code, message, tuple(sorted(where.items(), key=lambda item: item[0])))
        if key in self._seen:
            return
        self._seen.add(key)
        if len(self._warnings) >= self._limit:
            self._suppressed += 1
            return
        self._warnings.append(DocumentWarning(
            code=code,
            message=message,
            severity=self._severities.get(code, SEVERITY_WARNING),
            location=where,
        ))

    def records(self) -> List[DocumentWarning]:
        """Everything to report, including the truncation notice if any."""
        if not self._suppressed:
            return list(self._warnings)
        return self._warnings + [DocumentWarning(
            code=WARNINGS_TRUNCATED,
            message=(
                f"{self._suppressed} further warning(s) were not reported; fix "
                f"the ones above and rerun to see the rest."
            ),
            severity=SEVERITY_INFO,
        )]

    def as_dicts(self) -> List[Dict[str, Any]]:
        """The list ``main.py`` puts in the response."""
        return [warning.as_dict() for warning in self.records()]

    @property
    def messages(self) -> List[str]:
        """The rendered lines, for logging."""
        return [str(warning) for warning in self.records()]

    def __iter__(self) -> Iterator[DocumentWarning]:
        return iter(self.records())

    def __len__(self) -> int:
        return len(self.records())

    def __bool__(self) -> bool:
        return bool(self._warnings)

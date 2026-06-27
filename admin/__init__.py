"""Optional FastHTML template-admin UI and its supporting storage layer.

This package is **opt-in** (gated by ``ADMIN_ENABLED``) and is imported lazily
so that a deployment which never enables the admin UI pays no import cost and
needs none of the admin-only dependencies (e.g. python-fasthtml).

Phase 1 ships only the storage layer (:mod:`admin.store`); the FastHTML views
arrive in later phases.
"""
from __future__ import annotations

__all__ = ["store"]

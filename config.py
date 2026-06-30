"""Centralized configuration and logging setup for the MCP Office Documents server.

This module is the single source of truth for reading and validating environment
variables. No other module should access os.environ directly.

Highlights:
- Reads all env vars and constructs a typed Config instance (Pydantic v2).
- Validates required settings based on chosen upload strategy (LOCAL/S3/GCS/AZURE).
- Configures global logging (format and level) exactly once on first access.
- Exposes get_config() to retrieve a singleton Config across the app.

Environment variables (see .env.example for full list):
- Logging: DEBUG (true/false)
- Storage generic: UPLOAD_STRATEGY, SIGNED_URL_EXPIRES_IN
- Strategy specific: AWS_*, GCS_*, AZURE_*
"""

from __future__ import annotations

import logging
import os
from enum import Enum
from typing import Optional
from pydantic import BaseModel, Field, ValidationError, model_validator
from dotenv import load_dotenv

# Load .env file if present (local development). Existing env vars are NOT
# overwritten, so in Docker (where env_file sets them) this is a safe no-op.
load_dotenv()


class LogLevel(str, Enum):
    """Application log levels (restricted to INFO and DEBUG)."""
    DEBUG = "DEBUG"
    INFO = "INFO"


class LoggingSettings(BaseModel):
    """Logging configuration simplified to a single DEBUG flag.

    Behavior:
    - If DEBUG env var is truthy (1/true/on), logging is DEBUG.
    - Otherwise logging is INFO.

    Exposes convenience properties used across the app:
    - level_no: numeric logging level
    - mcp_level_str: lower-case string for FastMCP's `log_level` argument
    """
    debug: bool = Field(default=False, description="True to enable DEBUG level, False for INFO")

    @property
    def level_no(self) -> int:
        return logging.DEBUG if self.debug else logging.INFO

    @property
    def mcp_level_str(self) -> str:
        """Return lower-case string for FastMCP `log_level` argument."""
        return "debug" if self.debug else "info"


class S3Settings(BaseModel):
    """Configuration for AWS S3 uploads.

    Credential resolution follows a layered approach:

    1. **Explicit credentials** — Set ``AWS_ACCESS_KEY`` and
       ``AWS_SECRET_ACCESS_KEY`` environment variables. Can be used for local
       development or non-AWS environments. ``AWS_REGION`` is also required
       when using explicit credentials.
    2. **AWS default credential chain** — When
       the explicit credential env vars are *NOT* set, the boto3 SDK
       automatically discovers credentials from (in order):
       - Environment variables (``AWS_ACCESS_KEY_ID``, ``AWS_SECRET_ACCESS_KEY``,
         ``AWS_SESSION_TOKEN``, ``AWS_DEFAULT_REGION``)
       - Shared credential / config files (``~/.aws/credentials``,
         ``~/.aws/config``)
       - **AWS SSO / ``aws sso login`` sessions** — for local
         development after running ``aws sso login``
       - ECS container credentials
       - **IRSA (IAM Roles for Service Accounts)** — for pods running on
         **AWS EKS**
       - EC2 instance metadata (IMDSv2)

    ``S3_BUCKET`` is always required regardless of the credential method.
    ``AWS_REGION`` is optional when using the default credential chain — boto3
    will resolve the region from the environment, config files, or instance
    metadata.
    """
    access_key: Optional[str] = None
    secret_key: Optional[str] = None
    region: Optional[str] = None
    bucket: str

    @model_validator(mode="after")
    def _validate(self) -> "S3Settings":
        """Validate S3 settings.

        - ``bucket`` is always required.
        - If explicit credentials are partially provided (only one of
          access_key / secret_key), raise an error — either provide both
          or neither.
        - When explicit credentials are provided, ``region`` is required
          because we construct the endpoint URL from it.
        """
        # Normalize optional strings by stripping surrounding whitespace and
        # converting empty strings to None.
        if self.access_key is not None:
            self.access_key = self.access_key.strip() or None
        if self.secret_key is not None:
            self.secret_key = self.secret_key.strip() or None
        if self.region is not None:
            self.region = self.region.strip() or None

        if not self.bucket or not self.bucket.strip():
            raise ValueError("Missing required S3 setting: S3_BUCKET")
        self.bucket = self.bucket.strip()

        has_key = self.access_key is not None
        has_secret = self.secret_key is not None
        if has_key != has_secret:
            raise ValueError(
                "AWS_ACCESS_KEY and AWS_SECRET_ACCESS_KEY must both be set or both be omitted. "
                "Omit both to use the AWS default credential chain (IRSA, SSO, instance profile, etc.)."
            )

        if has_key and has_secret and self.region is None:
            raise ValueError(
                "AWS_REGION is required when explicit AWS_ACCESS_KEY / AWS_SECRET_ACCESS_KEY are provided."
            )

        return self

    @property
    def use_explicit_credentials(self) -> bool:
        """Return True if explicit AWS credentials were provided."""
        return self.access_key is not None and self.secret_key is not None


class GCSSettings(BaseModel):
    """Required configuration for Google Cloud Storage uploads.

    When credentials_path is omitted, the client uses Application Default
    Credentials (ADC), which automatically picks up Workload Identity on GKE.
    """
    bucket: str
    credentials_path: Optional[str] = None

    @model_validator(mode="after")
    def _non_empty(self) -> "GCSSettings":
        """Normalize credentials_path and ensure the bucket name is non-empty."""
        # Normalize credentials_path: treat whitespace-only strings as missing.
        if self.credentials_path is not None:
            stripped = str(self.credentials_path).strip()
            self.credentials_path = stripped or None

        # Validate and normalize bucket.
        bucket_stripped = str(self.bucket).strip()
        if not bucket_stripped:
            raise ValueError("Missing required GCS setting: GCS_BUCKET")
        self.bucket = bucket_stripped
        return self


class AzureSettings(BaseModel):
    """Required configuration for Azure Blob Storage uploads.

    Note: `endpoint` is optional; if empty, defaults to
    https://<account>.blob.core.windows.net
    """
    account_name: str
    account_key: str
    container: str
    endpoint: Optional[str] = None

    @model_validator(mode="after")
    def _non_empty(self) -> "AzureSettings":
        """Ensure all required Azure fields are non-empty."""
        missing = [
            name for name, val in (
                ("AZURE_STORAGE_ACCOUNT_NAME", self.account_name),
                ("AZURE_STORAGE_ACCOUNT_KEY", self.account_key),
                ("AZURE_CONTAINER", self.container),
            ) if not str(val).strip()
        ]
        if missing:
            raise ValueError(f"Missing required Azure settings: {', '.join(missing)}")
        return self


class MinioSettings(BaseModel):
    """Configuration for self-hosted MinIO (S3-compatible) uploads."""

    endpoint: str = Field(description="Base URL of the MinIO server, e.g., http://minio:9000")
    access_key: str
    secret_key: str
    bucket: str
    region: str = Field(default="us-east-1", description="Region to report to boto3; defaults to us-east-1")
    verify_ssl: bool = Field(default=True, description="Whether to verify SSL certificates when connecting")
    path_style: bool = Field(default=True, description="Use path-style addressing (recommended for MinIO)")

    @model_validator(mode="after")
    def _non_empty(self) -> "MinioSettings":
        missing = [
            name for name, val in (
                ("MINIO_ENDPOINT", self.endpoint),
                ("MINIO_ACCESS_KEY", self.access_key),
                ("MINIO_SECRET_KEY", self.secret_key),
                ("MINIO_BUCKET", self.bucket),
            ) if not str(val).strip()
        ]
        if missing:
            raise ValueError(f"Missing required MinIO settings: {', '.join(missing)}")
        return self


class StorageStrategy(str, Enum):
    """Supported upload backends for produced documents."""
    LOCAL = "LOCAL"
    S3 = "S3"
    GCS = "GCS"
    AZURE = "AZURE"
    MINIO = "MINIO"


class StorageSettings(BaseModel):
    """Generic storage configuration plus strategy-specific nested settings.

    Note: The LOCAL strategy always writes to the working folder ./app/upload;
    there is no configurable output directory for LOCAL.
    """
    strategy: StorageStrategy = Field(default=StorageStrategy.LOCAL)
    signed_url_expires_in: int = Field(default=3600, gt=0, description="TTL for S3/GCS/Azure download links in seconds")

    # Optional nested settings depending on strategy
    s3: Optional[S3Settings] = None
    gcs: Optional[GCSSettings] = None
    azure: Optional[AzureSettings] = None
    minio: Optional[MinioSettings] = None

    @model_validator(mode="after")
    def validate_strategy_requirements(self) -> "StorageSettings":
        """Ensure required nested settings exist for chosen strategy."""
        if self.strategy == StorageStrategy.S3:
            if not self.s3:
                raise ValueError("S3 settings are required for S3 strategy")
        elif self.strategy == StorageStrategy.GCS:
            if not self.gcs:
                raise ValueError("GCS settings are required for GCS strategy")
        elif self.strategy == StorageStrategy.AZURE:
            if not self.azure:
                raise ValueError("Azure settings are required for AZURE strategy")
        elif self.strategy == StorageStrategy.MINIO:
            if not self.minio:
                raise ValueError("MinIO settings are required for MINIO strategy")
        return self


class AdminSettings(BaseModel):
    """Configuration for the optional FastHTML template-admin UI.

    The admin UI is **opt-in**: when ``enabled`` is False (the default) the
    server behaves exactly as before — no admin routes are mounted and no
    extra dependencies are required at runtime. When enabled, a single shared
    password (``password``) gates the admin subtree mounted at ``path``.
    """
    enabled: bool = Field(default=False, description="Mount the template-admin UI when True.")
    password: Optional[str] = Field(
        default=None,
        description="Shared password gating the admin UI. Falls back to the API key when unset.",
    )
    path: str = Field(default="/admin", description="URL prefix the admin UI is mounted under.")

    @model_validator(mode="after")
    def _normalize(self) -> "AdminSettings":
        # Normalise the mount path to a leading-slash, no-trailing-slash form.
        p = (self.path or "/admin").strip()
        if not p.startswith("/"):
            p = "/" + p
        p = p.rstrip("/") or "/admin"
        self.path = p
        if self.password is not None:
            self.password = self.password.strip() or None
        return self


class Config(BaseModel):
    """Top-level configuration container used by the whole application."""
    logging: LoggingSettings
    storage: StorageSettings
    api_key: Optional[str] = Field(
        default=None,
        description="API key for authenticating incoming requests. None means no auth.",
    )
    admin: AdminSettings = Field(
        default_factory=AdminSettings,
        description="Settings for the optional template-admin UI.",
    )
    run_blocking_by_asyncio_thread_enabled: bool = Field(
        default=True,
        description=(
            "When True (default), MCP tool handlers dispatch their "
            "synchronous document-generation work to a bounded worker "
            "thread pool, keeping the FastMCP event loop free to serve "
            "health probes and concurrent requests. When False, tools "
            "run inline on the event loop — the original blocking "
            "behaviour (useful only for local debugging or to confirm a "
            "regression is caused by threading). Toggled via the "
            "RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED environment variable."
        ),
    )
    run_blocking_max_workers: int = Field(
        default=4,
        gt=0,
        description=(
            "Maximum concurrent worker threads in the bounded "
            "ThreadPoolExecutor used by run_blocking(). Lower values "
            "reduce GIL contention for CPU-bound work (markdown parsing); "
            "higher values allow more concurrent I/O-bound work (S3 "
            "upload). The default of 4 is tuned for 1 vCPU EKS pods: one "
            "thread doing CPU-bound parsing under the GIL while up to "
            "three others handle I/O. Tune via the "
            "RUN_BLOCKING_MAX_WORKERS environment variable."
        ),
    )
    stateless_http: bool = Field(
        default=False,
        description=(
            "When True, the FastMCP streamable-http transport runs in "
            "stateless mode: each request is independent, no Mcp-Session-Id "
            "is tracked, and any pod can serve any follow-up call. Required "
            "for horizontally-scaled EKS deployments (replicas >= 2) where "
            "the default in-process session dict would otherwise cause "
            "-32600 'Session not found' errors when the initialize and "
            "tools/call requests land on different pods. Default False "
            "preserves the stateful behaviour; this is safe for single-pod "
            "deployments and is what every existing deployment uses today. "
            "Toggled via the STATELESS_HTTP environment variable."
        ),
    )

    @property
    def admin_password_effective(self) -> Optional[str]:
        """Password the admin UI should accept.

        Prefers the explicit ``ADMIN_PASSWORD``; falls back to ``API_KEY`` so a
        single secret can protect both the MCP endpoint and the admin UI.
        Returns ``None`` when neither is set (admin auth cannot be satisfied).
        """
        return self.admin.password or self.api_key

    @staticmethod
    def _parse_bool(value: Optional[str]) -> bool:
        """Interpret common truthy representations used in env vars."""
        if value is None:
            return False
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}

    @classmethod
    def from_env(cls) -> "Config":
        """Construct Config from environment variables with sensible defaults and validation.

        This does not configure logging by itself; see configure_logging().
        """
        # Logging: only use DEBUG env var (truthy -> DEBUG, falsy -> INFO)
        debug = cls._parse_bool(os.environ.get("DEBUG"))
        logging_settings = LoggingSettings(debug=debug)

        # Storage
        raw_strategy = (os.environ.get("UPLOAD_STRATEGY", "LOCAL")).upper()
        strategy = raw_strategy if raw_strategy in {e.value for e in StorageStrategy} else "LOCAL"

        # Signed URL expiry (fallback to 3600 on invalid input)
        try:
            expires_in = int(os.environ.get("SIGNED_URL_EXPIRES_IN", "3600"))
            if expires_in <= 0:
                raise ValueError
        except ValueError:
            expires_in = 3600

        # Strategy-specific settings (only populate the relevant one)
        s3_settings = None
        gcs_settings = None
        azure_settings = None
        minio_settings = None

        if strategy == StorageStrategy.S3.value:
            s3_settings = S3Settings(
                access_key=os.environ.get("AWS_ACCESS_KEY") or None,
                secret_key=os.environ.get("AWS_SECRET_ACCESS_KEY") or None,
                region=os.environ.get("AWS_REGION") or None,
                bucket=os.environ.get("S3_BUCKET", ""),
            )
        elif strategy == StorageStrategy.GCS.value:
            gcs_settings = GCSSettings(
                bucket=os.environ.get("GCS_BUCKET", ""),
                credentials_path=os.environ.get("GCS_CREDENTIALS_PATH") or None,
            )
        elif strategy == StorageStrategy.AZURE.value:
            azure_settings = AzureSettings(
                account_name=os.environ.get("AZURE_STORAGE_ACCOUNT_NAME", ""),
                account_key=os.environ.get("AZURE_STORAGE_ACCOUNT_KEY", ""),
                container=os.environ.get("AZURE_CONTAINER", ""),
                endpoint=os.environ.get("AZURE_BLOB_ENDPOINT"),
            )
        elif strategy == StorageStrategy.MINIO.value:
            minio_settings = MinioSettings(
                endpoint=os.environ.get("MINIO_ENDPOINT", ""),
                access_key=os.environ.get("MINIO_ACCESS_KEY", ""),
                secret_key=os.environ.get("MINIO_SECRET_KEY", ""),
                bucket=os.environ.get("MINIO_BUCKET", ""),
                region=os.environ.get("MINIO_REGION", "us-east-1") or "us-east-1",
                verify_ssl=cls._parse_bool(os.environ.get("MINIO_VERIFY_SSL", "true")),
                path_style=cls._parse_bool(os.environ.get("MINIO_PATH_STYLE", "true")),
            )

        storage_settings = StorageSettings(
            strategy=StorageStrategy(strategy),
            signed_url_expires_in=expires_in,
            s3=s3_settings,
            gcs=gcs_settings,
            azure=azure_settings,
            minio=minio_settings,
        )

        # API key authentication (optional – empty/missing means no auth)
        raw_api_key = (os.environ.get("API_KEY") or "").strip() or None

        # Template-admin UI (optional, opt-in). Disabled unless ADMIN_ENABLED is
        # truthy. The admin password gates the UI; when unset it falls back to
        # the API key (so a single secret can protect both /mcp and /admin).
        admin_settings = AdminSettings(
            enabled=cls._parse_bool(os.environ.get("ADMIN_ENABLED")),
            password=(os.environ.get("ADMIN_PASSWORD") or "").strip() or None,
            path=(os.environ.get("ADMIN_PATH") or "/admin"),
        )

        # Toggle for thread-pool offload of blocking tool work. Default
        # True — when the env var is absent or empty we treat it as
        # "unset, use default". Set the var to "false"/"0"/"no"/"off" to
        # explicitly disable (e.g. for local debugging).
        raw_offload = os.environ.get("RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED")
        if raw_offload is None or not raw_offload.strip():
            run_blocking_by_asyncio_thread_enabled = True
        else:
            run_blocking_by_asyncio_thread_enabled = cls._parse_bool(raw_offload)

        # Bounded executor size. Invalid / non-positive values silently
        # fall back to the default (4) — tuned for 1 vCPU EKS pods.
        raw_workers = os.environ.get("RUN_BLOCKING_MAX_WORKERS")
        try:
            run_blocking_max_workers = int(raw_workers) if raw_workers is not None and raw_workers.strip() else 4
            if run_blocking_max_workers < 1:
                run_blocking_max_workers = 4
        except (ValueError, TypeError):
            run_blocking_max_workers = 4

        # Stateless streamable-http transport. Absent or empty env var
        # means False (existing behaviour). Truthy ("1"/"true"/"yes"/"on")
        # enables stateless mode for horizontal scaling.
        stateless_http = cls._parse_bool(os.environ.get("STATELESS_HTTP"))

        try:
            return cls(
                logging=logging_settings,
                storage=storage_settings,
                api_key=raw_api_key,
                admin=admin_settings,
                run_blocking_by_asyncio_thread_enabled=run_blocking_by_asyncio_thread_enabled,
                run_blocking_max_workers=run_blocking_max_workers,
                stateless_http=stateless_http,
            )
        except ValidationError as e:
            # Wrap Pydantic validation errors in a simpler exception for callers
            raise ValueError(f"Invalid configuration: {e}")


# Singleton instance and guard for one-time logging configuration
_CONFIG: Optional[Config] = None
_LOGGING_CONFIGURED: bool = False


def configure_logging(config: Config) -> None:
    """Configure root logger format and level exactly once.

    - Uses a more verbose format (file:line) in DEBUG level.
    - Keeps concise formatting otherwise.
    """
    global _LOGGING_CONFIGURED
    if _LOGGING_CONFIGURED:
        return

    level = config.logging.level_no
    root = logging.getLogger()
    if not root.handlers:
        handler = logging.StreamHandler()
        # Use the debug flag (simplified API)
        if config.logging.debug:
            fmt = "%(asctime)s | %(levelname)s | %(name)s:%(lineno)d | %(message)s"
        else:
            fmt = "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
        handler.setFormatter(logging.Formatter(fmt))
        root.addHandler(handler)
    root.setLevel(level)
    _LOGGING_CONFIGURED = True


def get_config() -> Config:
    """Return the process-wide Config singleton and ensure logging is configured."""
    global _CONFIG
    if _CONFIG is None:
        cfg = Config.from_env()
        configure_logging(cfg)
        _CONFIG = cfg
    return _CONFIG

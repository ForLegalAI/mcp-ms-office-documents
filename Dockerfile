# syntax=docker/dockerfile:1

# Comments are provided throughout this file to help you get started.
# If you need more help, visit the Dockerfile reference guide at
# https://docs.docker.com/go/dockerfile-reference/

# Roll with the latest 3.12 patch so the base ships current Alpine packages.
# Pin by digest here if you need fully reproducible builds.
ARG PYTHON_VERSION=3.12

# =============================================================================
# Stage 1: Builder - Install dependencies
# =============================================================================
FROM python:${PYTHON_VERSION}-alpine AS builder

# Patch base OS packages (openssl, musl, xz, sqlite, …) to the latest available.
RUN apk upgrade --no-cache

# Prevents Python from writing pyc files.
ENV PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# Install dependencies into a virtual environment for easy copying
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Download dependencies as a separate step to take advantage of Docker's caching.
# Leverage a cache mount to /root/.cache/pip to speed up subsequent builds.
# Leverage a bind mount to requirements.txt to avoid having to copy them into
# into this layer.
RUN --mount=type=cache,target=/root/.cache/pip \
    --mount=type=bind,source=requirements.txt,target=requirements.txt \
    pip install --no-cache-dir -r requirements.txt

# =============================================================================
# Stage 2: Runtime - Final lean image
# =============================================================================
FROM python:${PYTHON_VERSION}-alpine AS runtime

# Patch base OS packages in the SHIPPED image. This is the stage that matters for
# the scanned CVEs (openssl/musl/xz/sqlite); only /opt/venv is copied from builder.
RUN apk upgrade --no-cache

# Font files for the PowerPoint fit estimate, which measures text against a
# real face rather than counting characters. Each package is metric-compatible
# with a font corporate templates actually use, so measuring it measures the
# real thing: Carlito with Calibri, Liberation with Arial, Times New Roman and
# Courier New. Cambria and Georgia have no substitute here: Alpine packages
# neither Caladea nor Gelasio in main or community, so a deck in those faces is
# measured with a generic sans — a real measurement of the wrong font, the tier
# text_metrics documents. Naming a package Alpine does not carry fails the
# BUILD, not a test, which is how v4.0-beta.4 published no image at all;
# test_every_font_package_exists_in_alpine (network) checks these names against
# the package index, and test_every_metric_compatible_face_is_installed keeps
# this line in step with text_metrics.METRIC_COMPATIBLE. Without any font file
# the estimate falls back to arithmetic, which is what this replaces.
RUN apk add --no-cache font-carlito font-liberation

# Prevents Python from writing pyc files.
ENV PYTHONDONTWRITEBYTECODE=1

# Keeps Python from buffering stdout and stderr to avoid situations where
# the application crashes without emitting any logs due to buffering.
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Copy the virtual environment from builder stage
COPY --from=builder /opt/venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Create a non-privileged user that the app will run under.
# See https://docs.docker.com/go/dockerfile-user-best-practices/
ARG UID=10001
RUN adduser \
    --disabled-password \
    --gecos "" \
    --home "/nonexistent" \
    --shell "/sbin/nologin" \
    --no-create-home \
    --uid "${UID}" \
    appuser

# Create directories for output, custom templates, and config
RUN mkdir -p output custom_templates config

# Copy the source code into the container.
COPY . .

# Change ownership of directories to appuser
RUN chown -R appuser:appuser /app/output /app/custom_templates /app/config

# Switch to the non-privileged user to run the application.
USER appuser

# Expose the port that the application listens on.
EXPOSE 8958

# Run the application.
CMD ["python", "/app/main.py"]

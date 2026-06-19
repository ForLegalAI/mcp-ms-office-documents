# syntax=docker/dockerfile:1

# Comments are provided throughout this file to help you get started.
# If you need more help, visit the Dockerfile reference guide at
# https://docs.docker.com/go/dockerfile-reference/

ARG PYTHON_VERSION=3.12.8

# =============================================================================
# Stage 1: Builder - Install dependencies
# =============================================================================
# Using -slim (Debian/glibc) rather than -alpine (musl): the `formulas`
# library pulls in numpy and scipy, which publish pre-built wheels only
# for glibc. On Alpine these compile from source and need ~400MB of
# build tools (gcc, gfortran, musl-dev, linux-headers); on slim the
# wheels install directly with no build step, yielding faster builds
# and a smaller final image.
FROM python:${PYTHON_VERSION}-slim AS builder

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
FROM python:${PYTHON_VERSION}-slim AS runtime

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
RUN groupadd --system --gid "${UID}" appuser \
    && useradd --system --uid "${UID}" --gid appuser \
       --home-dir /nonexistent --shell /usr/sbin/nologin appuser

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

FROM ghcr.io/astral-sh/uv:latest

COPY . /app
WORKDIR /app

RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --no-dev

CMD ["uv", "run", "--no-sync", "vanlav_krd_bot/main.py"]

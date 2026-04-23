FROM ghcr.io/astral-sh/uv:latest

COPY . /app
WORKDIR /app

CMD ["uv", "run", "--no-dev", "vanlav_krd_bot/main.py"]

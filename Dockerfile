FROM ghcr.io/astral-sh/uv::python3.13-trixie

COPY . /app
WORKDIR /app

ENTRYPOINT ["uv", "run", "--no-dev", "vanlav_krd_bot/main.py"]

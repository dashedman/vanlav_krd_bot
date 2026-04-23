#!/usr/bin/env -S uv run

import asyncio
import logging
import sys

from bot import get_bot
from handlers import dp


async def start_listen():
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    bot_instance = get_bot()
    await dp.start_polling(bot_instance)


async def main():
    await start_listen()


if __name__ == "__main__":
    asyncio.run(main())

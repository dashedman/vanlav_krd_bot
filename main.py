import asyncio

import vanlav_krd_bot


async def main():
    await vanlav_krd_bot.start_listen()


if __name__ == "__main__":
    asyncio.run(main())

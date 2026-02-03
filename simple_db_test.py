import asyncpg
import asyncio

async def test_connection():
    try:
        conn = await asyncpg.connect(
            host='localhost',
            port=5433,
            user='postgres',
            password='postgres',
            database='pdfcon'
        )
        result = await conn.fetchval('SELECT 1')
        print(f"✅ Connection successful! Result: {result}")
        await conn.close()
        return True
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False

if __name__ == "__main__":
    asyncio.run(test_connection())
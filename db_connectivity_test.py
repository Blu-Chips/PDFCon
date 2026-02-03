#!/usr/bin/env python3
"""
Database Connectivity Test
Specifically tests the database connection with various approaches
"""

import asyncio
import asyncpg
import sqlalchemy
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession
from sqlalchemy.orm import sessionmaker
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

async def test_direct_asyncpg():
    """Test direct asyncpg connection"""
    print("🧪 Testing Direct AsyncPG Connection...")
    try:
        conn = await asyncpg.connect(
            host='localhost',
            port=5432,
            user='postgres',
            password='postgres',
            database='pdfcon'
        )
        version = await conn.fetchval('SELECT version()')
        await conn.close()
        print(f"✅ Direct AsyncPG: SUCCESS - {version[:50]}...")
        return True
    except Exception as e:
        print(f"❌ Direct AsyncPG: FAILED - {str(e)}")
        return False

async def test_sqlalchemy_async():
    """Test SQLAlchemy async connection"""
    print("\n🧪 Testing SQLAlchemy Async Connection...")
    try:
        DATABASE_URL = "postgresql+asyncpg://postgres:postgres@localhost:5432/pdfcon"
        engine = create_async_engine(DATABASE_URL, echo=False)
        
        async with engine.connect() as conn:
            result = await conn.execute(sqlalchemy.text("SELECT version()"))
            version = result.scalar()
            print(f"✅ SQLAlchemy Async: SUCCESS - {version[:50]}...")
            await engine.dispose()
            return True
    except Exception as e:
        print(f"❌ SQLAlchemy Async: FAILED - {str(e)}")
        return False

async def test_env_config_connection():
    """Test connection using environment configuration"""
    print("\n🧪 Testing Environment Config Connection...")
    try:
        from backend.app.core.database import engine
        
        async with engine.connect() as conn:
            result = await conn.execute(sqlalchemy.text("SELECT 1"))
            value = result.scalar()
            print(f"✅ Environment Config: SUCCESS - Connected (test value: {value})")
            return True
    except Exception as e:
        print(f"❌ Environment Config: FAILED - {str(e)}")
        return False

async def test_table_creation():
    """Test table creation capability"""
    print("\n🧪 Testing Table Creation...")
    try:
        from backend.app.core.database import engine, Base
        from backend.app.core.config import settings
        
        # Create a simple test table
        from sqlalchemy import Column, Integer, String, DateTime
        from sqlalchemy.ext.declarative import declarative_base
        
        class TestModel(Base):
            __tablename__ = 'test_table'
            id = Column(Integer, primary_key=True)
            name = Column(String(50))
            created_at = Column(DateTime)
        
        async with engine.begin() as conn:
            await conn.run_sync(Base.metadata.create_all)
            print("✅ Table Creation: SUCCESS - Tables created")
            return True
    except Exception as e:
        print(f"❌ Table Creation: FAILED - {str(e)}")
        return False

async def main():
    """Run all database connectivity tests"""
    print("=" * 60)
    print("🔍 DATABASE CONNECTIVITY CLINICAL TEST")
    print("=" * 60)
    
    tests = [
        test_direct_asyncpg,
        test_sqlalchemy_async,
        test_env_config_connection,
        test_table_creation
    ]
    
    results = []
    for test in tests:
        try:
            result = await test()
            results.append(result)
        except Exception as e:
            print(f"💥 Test crashed: {str(e)}")
            results.append(False)
    
    # Summary
    passed = sum(results)
    total = len(results)
    success_rate = (passed / total) * 100 if total > 0 else 0
    
    print("\n" + "=" * 60)
    print("📊 DATABASE TEST SUMMARY")
    print("=" * 60)
    print(f"Tests Passed: {passed}/{total}")
    print(f"Success Rate: {success_rate:.1f}%")
    
    if success_rate >= 75:
        print("🎉 Database connectivity is CLINICALLY HEALTHY!")
        return 0
    elif success_rate >= 50:
        print("⚠️ Database connectivity has MINOR ISSUES but is mostly functional")
        return 1
    else:
        print("💥 Database connectivity has CRITICAL ISSUES")
        return 2

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    exit(exit_code)
#!/usr/bin/env python3
"""
Simple Performance Test
Direct testing without complex imports
"""

import asyncio
import time
import sys
import os

# Add backend to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

async def simple_performance_test():
    print("=" * 60)
    print("⚡ SIMPLE PERFORMANCE TEST")
    print("=" * 60)
    
    # Test 1: Basic import timing
    print("\n1. Testing Import Performance...")
    start_time = time.time()
    
    try:
        from backend.app.agents.orchestrator import OrchestratorAgent
        from backend.app.agents.scraping_agent import ScrapingAgent
        from backend.app.agents.processing_agent import ProcessingAgent
        
        import_time = time.time() - start_time
        print(f"✅ Imports successful in {import_time:.4f} seconds")
        
    except Exception as e:
        print(f"❌ Import failed: {e}")
        return False
    
    # Test 2: Agent initialization timing
    print("\n2. Testing Agent Initialization Timing...")
    init_times = []
    
    for i in range(10):
        start_time = time.time()
        try:
            agent = OrchestratorAgent()
            elapsed = time.time() - start_time
            init_times.append(elapsed)
        except Exception as e:
            print(f"❌ Agent init failed on iteration {i}: {e}")
            return False
    
    avg_init_time = sum(init_times) / len(init_times)
    min_init_time = min(init_times)
    max_init_time = max(init_times)
    
    print(f"✅ Agent initialization:")
    print(f"   → Average: {avg_init_time*1000:.2f} ms")
    print(f"   → Min: {min_init_time*1000:.2f} ms")
    print(f"   → Max: {max_init_time*1000:.2f} ms")
    
    # Test 3: URL validation performance
    print("\n3. Testing URL Validation Performance...")
    scraper = ScrapingAgent()
    
    test_urls = [
        "https://example.com",
        "http://test.org/page", 
        "https://www.google.com/search?q=test",
        "not-a-url",
        "ftp://invalid.com",
        ""
    ] * 200  # 1200 URLs
    
    start_time = time.time()
    valid_count = 0
    
    for url in test_urls:
        if scraper.validate_url(url):
            valid_count += 1
    
    validation_time = time.time() - start_time
    validation_rate = len(test_urls) / validation_time
    
    print(f"✅ URL Validation:")
    print(f"   → Processed: {len(test_urls)} URLs in {validation_time:.3f} seconds")
    print(f"   → Valid: {valid_count} URLs")
    print(f"   → Rate: {validation_rate:.2f} URLs/second")
    print(f"   → Per URL: {validation_time/len(test_urls)*1000:.3f} ms")
    
    # Test 4: Text processing performance
    print("\n4. Testing Text Processing Performance...")
    processor = ProcessingAgent()
    
    base_text = "  This is   TEST   text with EXTRA   spaces  \n\n  and newlines  "
    test_texts = [base_text * (i % 3 + 1) for i in range(500)]
    
    start_time = time.time()
    processed_texts = []
    
    for text in test_texts:
        cleaned = processor.clean_text(text)
        processed_texts.append(cleaned)
    
    processing_time = time.time() - start_time
    processing_rate = len(test_texts) / processing_time
    
    avg_reduction = sum(len(orig) - len(proc) for orig, proc in zip(test_texts, processed_texts)) / len(test_texts)
    
    print(f"✅ Text Processing:")
    print(f"   → Processed: {len(test_texts)} texts in {processing_time:.3f} seconds")
    print(f"   → Rate: {processing_rate:.2f} texts/second")
    print(f"   → Avg reduction: {avg_reduction:.1f} characters")
    print(f"   → Per text: {processing_time/len(test_texts)*1000:.3f} ms")
    
    # Test 5: Concurrent operations
    print("\n5. Testing Concurrent Operations...")
    
    async def concurrent_init():
        return OrchestratorAgent()
    
    start_time = time.time()
    tasks = [concurrent_init() for _ in range(15)]
    results = await asyncio.gather(*tasks, return_exceptions=True)
    concurrent_time = time.time() - start_time
    
    successful = sum(1 for r in results if not isinstance(r, Exception))
    concurrent_rate = successful / concurrent_time
    
    print(f"✅ Concurrent Operations:")
    print(f"   → {successful}/15 agents initialized in {concurrent_time:.3f} seconds")
    print(f"   → Rate: {concurrent_rate:.2f} initializations/second")
    
    # Summary
    print("\n" + "=" * 60)
    print("📈 PERFORMANCE SUMMARY")
    print("=" * 60)
    
    overall_rate = (len(test_urls) + len(test_texts) + successful) / (validation_time + processing_time + concurrent_time)
    
    print(f"Overall Processing Rate: {overall_rate:.2f} operations/second")
    
    if overall_rate > 500:
        rating = "🚀 EXCELLENT"
    elif overall_rate > 200:
        rating = "✅ GOOD"
    elif overall_rate > 100:
        rating = "⚠️  ACCEPTABLE"
    else:
        rating = "❌ NEEDS IMPROVEMENT"
    
    print(f"Performance Rating: {rating}")
    
    # Save results
    import json
    results_data = {
        "import_time": round(import_time, 4),
        "agent_init_avg_ms": round(avg_init_time * 1000, 2),
        "url_validation_rate": round(validation_rate, 2),
        "text_processing_rate": round(processing_rate, 2),
        "concurrent_rate": round(concurrent_rate, 2),
        "overall_rate": round(overall_rate, 2),
        "rating": rating,
        "timestamp": time.time()
    }
    
    with open('simple_performance_results.json', 'w') as f:
        json.dump(results_data, f, indent=2)
    
    print(f"\n💾 Results saved to: simple_performance_results.json")
    
    return overall_rate > 100  # Pass if reasonable performance

async def main():
    success = await simple_performance_test()
    return 0 if success else 1

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    exit(exit_code)
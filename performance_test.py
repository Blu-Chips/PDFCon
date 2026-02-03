#!/usr/bin/env python3
"""
GEDI Performance and Stress Testing
Tests system performance under various loads
"""

import asyncio
import time
import statistics
from concurrent.futures import ThreadPoolExecutor
import aiohttp
from typing import List, Dict, Any

class PerformanceTester:
    def __init__(self):
        self.results = {}
        self.metrics = {}
        
    def log_metric(self, test_name: str, metric: str, value: Any):
        """Log performance metrics"""
        if test_name not in self.metrics:
            self.metrics[test_name] = {}
        self.metrics[test_name][metric] = value
        
    async def test_concurrent_agent_initialization(self, concurrent_count: int = 10) -> Dict:
        """Test concurrent agent initialization performance"""
        print(f"\n🚀 Testing Concurrent Agent Initialization ({concurrent_count} concurrent)")
        print("-" * 60)
        
        start_time = time.time()
        
        async def init_agent():
            from backend.app.agents.orchestrator import OrchestratorAgent
            agent = OrchestratorAgent()
            return agent
            
        # Create concurrent tasks
        tasks = [init_agent() for _ in range(concurrent_count)]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        end_time = time.time()
        duration = end_time - start_time
        
        successful = sum(1 for r in results if not isinstance(r, Exception))
        failure_rate = (concurrent_count - successful) / concurrent_count * 100
        
        self.log_metric("Concurrent Agent Init", "concurrent_count", concurrent_count)
        self.log_metric("Concurrent Agent Init", "duration_seconds", round(duration, 3))
        self.log_metric("Concurrent Agent Init", "successful_initializations", successful)
        self.log_metric("Concurrent Agent Init", "failure_rate_percent", round(failure_rate, 2))
        self.log_metric("Concurrent Agent Init", "rate_per_second", round(successful/duration, 2))
        
        status = "✅ PASS" if failure_rate == 0 else "⚠️  PARTIAL" if failure_rate < 10 else "❌ FAIL"
        print(f"{status} Concurrent Agent Init: {successful}/{concurrent_count} successful in {duration:.3f}s")
        print(f"   → Rate: {successful/duration:.2f} initializations/second")
        print(f"   → Failure rate: {failure_rate:.1f}%")
        
        return {
            "test": "Concurrent Agent Init",
            "successful": successful,
            "total": concurrent_count,
            "duration": duration,
            "rate": successful/duration
        }
    
    async def test_url_validation_performance(self, url_count: int = 1000) -> Dict:
        """Test URL validation performance at scale"""
        print(f"\n🔍 Testing URL Validation Performance ({url_count} URLs)")
        print("-" * 60)
        
        from backend.app.agents.scraping_agent import ScrapingAgent
        agent = ScrapingAgent()
        
        # Generate test URLs
        test_urls = [
            "https://example.com",
            "http://test.org/page",
            "https://www.google.com/search?q=test",
            "not-a-url",
            "ftp://invalid.com",
            ""
        ] * (url_count // 6 + 1)
        test_urls = test_urls[:url_count]
        
        start_time = time.time()
        
        # Validate URLs
        results = []
        for url in test_urls:
            try:
                result = agent.validate_url(url)
                results.append(result)
            except Exception:
                results.append(False)
        
        end_time = time.time()
        duration = end_time - start_time
        
        valid_count = sum(1 for r in results if r)
        validation_rate = len(test_urls) / duration
        
        self.log_metric("URL Validation", "url_count", url_count)
        self.log_metric("URL Validation", "duration_seconds", round(duration, 3))
        self.log_metric("URL Validation", "valid_urls", valid_count)
        self.log_metric("URL Validation", "validation_rate", round(validation_rate, 2))
        
        print(f"✅ URL Validation: {valid_count}/{url_count} valid in {duration:.3f}s")
        print(f"   → Rate: {validation_rate:.2f} validations/second")
        print(f"   → Average: {duration/url_count*1000:.3f} ms per validation")
        
        return {
            "test": "URL Validation",
            "valid": valid_count,
            "total": url_count,
            "duration": duration,
            "rate": validation_rate
        }
    
    async def test_text_processing_performance(self, text_samples: int = 500) -> Dict:
        """Test text processing performance"""
        print(f"\n📝 Testing Text Processing Performance ({text_samples} samples)")
        print("-" * 60)
        
        from backend.app.agents.processing_agent import ProcessingAgent
        agent = ProcessingAgent()
        
        # Generate test texts
        base_text = "  This is   TEST   text with EXTRA   spaces  \n\n  and newlines  "
        test_texts = [base_text * (i % 5 + 1) for i in range(text_samples)]
        
        start_time = time.time()
        
        # Process texts
        results = []
        for text in test_texts:
            try:
                cleaned = agent.clean_text(text)
                results.append(len(cleaned))
            except Exception as e:
                results.append(0)
        
        end_time = time.time()
        duration = end_time - start_time
        
        successful = sum(1 for r in results if r > 0)
        processing_rate = successful / duration
        
        avg_length_reduction = statistics.mean([
            len(orig) - results[i] 
            for i, orig in enumerate(test_texts) 
            if results[i] > 0
        ]) if successful > 0 else 0
        
        self.log_metric("Text Processing", "samples", text_samples)
        self.log_metric("Text Processing", "duration_seconds", round(duration, 3))
        self.log_metric("Text Processing", "successful", successful)
        self.log_metric("Text Processing", "processing_rate", round(processing_rate, 2))
        self.log_metric("Text Processing", "avg_length_reduction", round(avg_length_reduction, 2))
        
        print(f"✅ Text Processing: {successful}/{text_samples} processed in {duration:.3f}s")
        print(f"   → Rate: {processing_rate:.2f} texts/second")
        print(f"   → Avg reduction: {avg_length_reduction:.1f} characters")
        print(f"   → Average: {duration/text_samples*1000:.3f} ms per text")
        
        return {
            "test": "Text Processing",
            "successful": successful,
            "total": text_samples,
            "duration": duration,
            "rate": processing_rate,
            "avg_reduction": avg_length_reduction
        }
    
    async def test_memory_usage_baseline(self) -> Dict:
        """Test baseline memory usage"""
        print("\n🧠 Testing Memory Usage Baseline")
        print("-" * 60)
        
        import psutil
        import os
        
        # Get process info
        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        
        rss_mb = memory_info.rss / 1024 / 1024
        vms_mb = memory_info.vms / 1024 / 1024
        
        self.log_metric("Memory Usage", "rss_mb", round(rss_mb, 2))
        self.log_metric("Memory Usage", "vms_mb", round(vms_mb, 2))
        
        status = "✅ LOW" if rss_mb < 100 else "⚠️  MODERATE" if rss_mb < 500 else "❌ HIGH"
        print(f"{status} Memory Usage: {rss_mb:.2f} MB RSS, {vms_mb:.2f} MB VMS")
        
        return {
            "test": "Memory Usage",
            "rss_mb": rss_mb,
            "vms_mb": vms_mb
        }
    
    async def run_performance_suite(self) -> Dict:
        """Run complete performance test suite"""
        print("=" * 70)
        print("⚡ GEDI PERFORMANCE AND STRESS TEST SUITE")
        print("=" * 70)
        
        test_results = []
        
        # Run performance tests
        agent_init_result = await self.test_concurrent_agent_initialization(20)
        test_results.append(agent_init_result)
        
        url_validation_result = await self.test_url_validation_performance(2000)
        test_results.append(url_validation_result)
        
        text_processing_result = await self.test_text_processing_performance(1000)
        test_results.append(text_processing_result)
        
        memory_result = await self.test_memory_usage_baseline()
        test_results.append(memory_result)
        
        # Generate summary
        print("\n" + "=" * 70)
        print("📈 PERFORMANCE TEST SUMMARY")
        print("=" * 70)
        
        total_duration = sum(r.get("duration", 0) for r in test_results)
        avg_rate = statistics.mean([r.get("rate", 0) for r in test_results if "rate" in r]) if test_results else 0
        
        print(f"Total Test Duration: {total_duration:.2f} seconds")
        print(f"Average Processing Rate: {avg_rate:.2f} operations/second")
        
        # Performance rating
        if avg_rate > 1000:
            rating = "🚀 EXCELLENT"
        elif avg_rate > 500:
            rating = "✅ GOOD"
        elif avg_rate > 100:
            rating = "⚠️  ACCEPTABLE"
        else:
            rating = "❌ POOR"
            
        print(f"Overall Performance Rating: {rating}")
        
        # Detailed metrics
        print("\n📋 DETAILED METRICS:")
        for test_name, metrics in self.metrics.items():
            print(f"\n{test_name}:")
            for metric, value in metrics.items():
                print(f"  {metric}: {value}")
        
        # Save results
        import json
        results_data = {
            "summary": {
                "total_duration": round(total_duration, 2),
                "average_rate": round(avg_rate, 2),
                "performance_rating": rating,
                "timestamp": time.time()
            },
            "test_results": test_results,
            "detailed_metrics": self.metrics
        }
        
        with open('performance_test_results.json', 'w') as f:
            json.dump(results_data, f, indent=2)
        
        print(f"\n💾 Results saved to: performance_test_results.json")
        
        return results_data

async def main():
    """Main performance test runner"""
    tester = PerformanceTester()
    results = await tester.run_performance_suite()
    
    # Return appropriate exit code
    rating = results["summary"]["performance_rating"]
    if "EXCELLENT" in rating or "GOOD" in rating:
        return 0
    elif "ACCEPTABLE" in rating:
        return 1
    else:
        return 2

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    exit(exit_code)
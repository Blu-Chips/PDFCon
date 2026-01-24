"""
Comparison Agent
Benchmarks against Norway Sovereign Wealth Fund
"""
import logging
from typing import Dict

logger = logging.getLogger(__name__)


class ComparisonAgent:
    """Agent responsible for comparative analysis"""
    
    def __init__(self):
        """Initialize the comparison agent"""
        logger.info("Comparison agent initialized")
    
    async def compare(self, analysis: Dict, year: int) -> Dict:
        """
        Compare analysis with Norway Sovereign Wealth Fund data
        
        Args:
            analysis: Financial analysis results
            year: Report year
            
        Returns:
            Dict containing comparative analysis
        """
        logger.info(f"Comparing with Norway Sovereign Wealth Fund for {year}")
        # TODO: Implement Norway benchmarking logic
        return {
            "performance_comparison": {},
            "risk_comparison": {},
            "allocation_comparison": {},
            "recommendations": []
        }
    
    async def health_check(self) -> str:
        """Check health of the agent"""
        return "healthy"
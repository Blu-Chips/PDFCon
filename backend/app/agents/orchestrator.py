"""
AI Orchestrator Agent
Coordinates all sub-agents for end-to-end report processing
"""
import asyncio
import logging
from typing import Dict, List, Optional
from datetime import datetime
from enum import Enum

from app.core.config import settings
from app.agents.scraping_agent import ScrapingAgent
from app.agents.processing_agent import ProcessingAgent
from app.agents.analysis_agent import AnalysisAgent
from app.agents.comparison_agent import ComparisonAgent
from app.agents.reporting_agent import ReportingAgent

logger = logging.getLogger(__name__)


class ProcessingStatus(Enum):
    """Processing status enumeration"""
    INITIATED = "initiated"
    SCRAPING = "scraping"
    PROCESSING = "processing"
    ANALYZING = "analyzing"
    COMPARING = "comparing"
    REPORTING = "reporting"
    COMPLETED = "completed"
    FAILED = "failed"


class OrchestratorAgent:
    """
    Primary Orchestrator Agent
    Coordinates all sub-agents and manages the complete workflow
    """
    
    def __init__(self):
        """Initialize the orchestrator with all sub-agents"""
        self.scraping_agent = ScrapingAgent()
        self.processing_agent = ProcessingAgent()
        self.analysis_agent = AnalysisAgent()
        self.comparison_agent = ComparisonAgent()
        self.reporting_agent = ReportingAgent()
        
        self.current_status: ProcessingStatus = ProcessingStatus.INITIATED
        self.start_time: Optional[datetime] = None
        self.progress: float = 0.0
        
        logger.info("Orchestrator agent initialized")
    
    async def process_report(
        self,
        country: str,
        year: int,
        source: str = "scraping",
        file_path: Optional[str] = None
    ) -> Dict:
        """
        Main orchestration workflow for processing a government financial report
        
        Args:
            country: Country or county name
            year: Report year
            source: 'scraping' or 'upload'
            file_path: Optional file path for uploaded documents
            
        Returns:
            Dict containing complete analysis results and report
        """
        self.start_time = datetime.now()
        logger.info(f"Starting report processing for {country}, year {year}")
        
        try:
            # Step 1: Collect documents
            self.current_status = ProcessingStatus.SCRAPING
            self.progress = 10.0
            logger.info("Step 1: Collecting documents...")
            
            if source == "upload" and file_path:
                documents = await self.scraping_agent.load_uploaded_file(file_path, country, year)
            else:
                documents = await self.scraping_agent.collect(country, year)
            
            if not documents:
                raise ValueError(f"No documents found for {country}, year {year}")
            
            logger.info(f"Collected {len(documents)} document(s)")
            
            # Step 2: Process documents
            self.current_status = ProcessingStatus.PROCESSING
            self.progress = 30.0
            logger.info("Step 2: Processing documents...")
            
            processed_data = await self.processing_agent.process(documents)
            
            if not processed_data:
                raise ValueError("Failed to process documents")
            
            logger.info("Documents processed successfully")
            
            # Step 3: Analyze data
            self.current_status = ProcessingStatus.ANALYZING
            self.progress = 50.0
            logger.info("Step 3: Analyzing financial data...")
            
            analysis = await self.analysis_agent.analyze(processed_data)
            
            if not analysis:
                raise ValueError("Failed to analyze data")
            
            logger.info("Financial analysis completed")
            
            # Step 4: Compare with Norway
            self.current_status = ProcessingStatus.COMPARING
            self.progress = 70.0
            logger.info("Step 4: Comparing with Norway Sovereign Wealth Fund...")
            
            comparison = await self.comparison_agent.compare(analysis, year)
            
            logger.info("Comparative analysis completed")
            
            # Step 5: Generate report
            self.current_status = ProcessingStatus.REPORTING
            self.progress = 90.0
            logger.info("Step 5: Generating comprehensive report...")
            
            report = await self.reporting_agent.generate(analysis, comparison)
            
            if not report:
                raise ValueError("Failed to generate report")
            
            logger.info("Report generated successfully")
            
            # Finalize
            self.current_status = ProcessingStatus.COMPLETED
            self.progress = 100.0
            
            end_time = datetime.now()
            duration = (end_time - self.start_time).total_seconds()
            
            result = {
                "status": "success",
                "country": country,
                "year": year,
                "processing_time_seconds": duration,
                "documents_processed": len(documents),
                "analysis": analysis,
                "comparison": comparison,
                "report": report,
                "metadata": {
                    "start_time": self.start_time.isoformat(),
                    "end_time": end_time.isoformat(),
                    "duration": duration,
                    "steps_completed": 5
                }
            }
            
            logger.info(f"Report processing completed in {duration:.2f} seconds")
            return result
            
        except Exception as e:
            self.current_status = ProcessingStatus.FAILED
            logger.error(f"Report processing failed: {str(e)}", exc_info=True)
            
            return {
                "status": "error",
                "country": country,
                "year": year,
                "error": str(e),
                "current_step": self.current_status.value,
                "progress": self.progress
            }
    
    async def get_status(self) -> Dict:
        """
        Get current processing status
        
        Returns:
            Dict containing current status and progress
        """
        return {
            "status": self.current_status.value,
            "progress": self.progress,
            "start_time": self.start_time.isoformat() if self.start_time else None,
            "elapsed_time": (datetime.now() - self.start_time).total_seconds() if self.start_time else 0
        }
    
    async def process_batch(
        self,
        reports: List[Dict]
    ) -> List[Dict]:
        """
        Process multiple reports in parallel
        
        Args:
            reports: List of dicts with 'country', 'year', 'source', 'file_path'
            
        Returns:
            List of processing results
        """
        logger.info(f"Processing batch of {len(reports)} reports")
        
        tasks = []
        for report in reports:
            task = self.process_report(
                country=report["country"],
                year=report["year"],
                source=report.get("source", "scraping"),
                file_path=report.get("file_path")
            )
            tasks.append(task)
        
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        successful = sum(1 for r in results if isinstance(r, dict) and r.get("status") == "success")
        failed = len(results) - successful
        
        logger.info(f"Batch processing completed: {successful} successful, {failed} failed")
        
        return results
    
    async def health_check(self) -> Dict:
        """
        Check health of all sub-agents
        
        Returns:
            Dict containing health status of all agents
        """
        health_status = {
            "orchestrator": "healthy",
            "sub_agents": {}
        }
        
        # Check each sub-agent
        try:
            health_status["sub_agents"]["scraping"] = await self.scraping_agent.health_check()
        except Exception as e:
            health_status["sub_agents"]["scraping"] = f"unhealthy: {str(e)}"
        
        try:
            health_status["sub_agents"]["processing"] = await self.processing_agent.health_check()
        except Exception as e:
            health_status["sub_agents"]["processing"] = f"unhealthy: {str(e)}"
        
        try:
            health_status["sub_agents"]["analysis"] = await self.analysis_agent.health_check()
        except Exception as e:
            health_status["sub_agents"]["analysis"] = f"unhealthy: {str(e)}"
        
        try:
            health_status["sub_agents"]["comparison"] = await self.comparison_agent.health_check()
        except Exception as e:
            health_status["sub_agents"]["comparison"] = f"unhealthy: {str(e)}"
        
        try:
            health_status["sub_agents"]["reporting"] = await self.reporting_agent.health_check()
        except Exception as e:
            health_status["sub_agents"]["reporting"] = f"unhealthy: {str(e)}"
        
        # Determine overall health
        all_healthy = all(
            "unhealthy" not in str(status) 
            for status in health_status["sub_agents"].values()
        )
        
        health_status["overall"] = "healthy" if all_healthy else "degraded"
        
        return health_status


# Create singleton instance
orchestrator = OrchestratorAgent()
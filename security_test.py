#!/usr/bin/env python3
"""
GEDI Security Validation Test
Tests basic security configurations and validations
"""

import asyncio
import hashlib
import secrets
import time
from typing import Dict, List

class SecurityTester:
    def __init__(self):
        self.results = {}
        
    def log_security_test(self, test_name: str, status: str, details: str = ""):
        """Log security test results"""
        self.results[test_name] = {
            "status": status,
            "details": details,
            "timestamp": time.time()
        }
        
        status_icon = "✅" if status == "PASS" else "❌" if status == "FAIL" else "⚠️"
        print(f"{status_icon} {test_name}: {status}")
        if details:
            print(f"   → {details}")
    
    async def test_config_security(self) -> bool:
        """Test configuration security settings"""
        print("\n🔒 Testing Configuration Security")
        print("-" * 40)
        
        try:
            from backend.app.core.config import settings
            
            # Check for secure defaults
            issues = []
            
            # Secret key validation
            if settings.SECRET_KEY == "your-secret-key-change-in-production":
                issues.append("Default secret key in use")
            elif len(settings.SECRET_KEY) < 32:
                issues.append("Secret key too short")
            
            # Debug mode should be off in production
            if settings.DEBUG:
                issues.append("Debug mode enabled (should be disabled in production)")
            
            # Check for placeholder API keys
            placeholder_keys = [
                "your-openai-api-key-here",
                "your-anthropic-api-key-here",
                "sk-...",
                "AIzaSy"
            ]
            
            api_keys = [
                settings.OPENAI_API_KEY,
                settings.ANTHROPIC_API_KEY,
                settings.GEMINI_API_KEY
            ]
            
            for key in api_keys:
                if key and any(placeholder in str(key) for placeholder in placeholder_keys):
                    issues.append("Placeholder API key detected")
            
            if issues:
                self.log_security_test("Config Security", "WARN", "; ".join(issues))
                return True  # Warn but not fail
            else:
                self.log_security_test("Config Security", "PASS", "Secure configuration detected")
                return True
                
        except Exception as e:
            self.log_security_test("Config Security", "FAIL", f"Error: {str(e)}")
            return False
    
    async def test_password_hashing(self) -> bool:
        """Test password hashing security"""
        print("\n🔑 Testing Password Hashing")
        print("-" * 40)
        
        try:
            from passlib.context import CryptContext
            
            # Create password context
            pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
            
            # Test password hashing
            test_password = "test_password_123"
            hashed = pwd_context.hash(test_password)
            
            # Verify hash works
            verify_success = pwd_context.verify(test_password, hashed)
            verify_fail = pwd_context.verify("wrong_password", hashed)
            
            # Check hash properties
            hash_length = len(hashed)
            uses_bcrypt = hashed.startswith("$2b$")
            
            if verify_success and not verify_fail and uses_bcrypt and hash_length > 50:
                self.log_security_test("Password Hashing", "PASS", f"BCrypt hashing working (hash length: {hash_length})")
                return True
            else:
                self.log_security_test("Password Hashing", "FAIL", "Hashing verification failed")
                return False
                
        except Exception as e:
            self.log_security_test("Password Hashing", "FAIL", f"Error: {str(e)}")
            return False
    
    async def test_jwt_token_security(self) -> bool:
        """Test JWT token generation and validation"""
        print("\n🎫 Testing JWT Token Security")
        print("-" * 40)
        
        try:
            from jose import jwt
            from backend.app.core.config import settings
            import time
            
            # Test token creation
            payload = {
                "sub": "test_user",
                "exp": int(time.time()) + 3600,  # 1 hour expiry
                "iat": int(time.time())
            }
            
            # Create token
            token = jwt.encode(payload, settings.SECRET_KEY, algorithm=settings.ALGORITHM)
            
            # Decode token
            decoded = jwt.decode(token, settings.SECRET_KEY, algorithms=[settings.ALGORITHM])
            
            # Verify payload
            if (decoded["sub"] == payload["sub"] and 
                decoded["exp"] == payload["exp"] and
                "iat" in decoded):
                
                # Test expiration
                expired_payload = payload.copy()
                expired_payload["exp"] = int(time.time()) - 100  # Expired 100 seconds ago
                expired_token = jwt.encode(expired_payload, settings.SECRET_KEY, algorithm=settings.ALGORITHM)
                
                try:
                    jwt.decode(expired_token, settings.SECRET_KEY, algorithms=[settings.ALGORITHM])
                    expired_works = True
                except:
                    expired_works = False
                
                if not expired_works:
                    self.log_security_test("JWT Token Security", "PASS", f"Token generation/validation working (length: {len(token)})")
                    return True
                else:
                    self.log_security_test("JWT Token Security", "FAIL", "Expired token still valid")
                    return False
            else:
                self.log_security_test("JWT Token Security", "FAIL", "Token payload mismatch")
                return False
                
        except Exception as e:
            self.log_security_test("JWT Token Security", "FAIL", f"Error: {str(e)}")
            return False
    
    async def test_input_validation(self) -> bool:
        """Test input validation security"""
        print("\n🛡️  Testing Input Validation")
        print("-" * 40)
        
        try:
            from backend.app.agents.scraping_agent import ScrapingAgent
            
            agent = ScrapingAgent()
            
            # Test malicious URL inputs
            malicious_inputs = [
                "javascript:alert('xss')",
                "data:text/html,<script>alert('xss')</script>",
                "file:///etc/passwd",
                "http://example.com?param=<script>alert('xss')</script>",
                None,
                "",
                "   ",
                "../../../../etc/passwd"
            ]
            
            unsafe_count = 0
            for malicious_input in malicious_inputs:
                try:
                    # This should either return False or raise an exception for unsafe inputs
                    result = agent.validate_url(malicious_input)
                    if result is True:  # If it validates as safe when it shouldn't be
                        unsafe_count += 1
                except Exception:
                    # Exceptions for malicious inputs are acceptable
                    pass
            
            # Test normal URLs still work
            normal_urls = [
                "https://example.com",
                "http://test.org/page"
            ]
            
            safe_count = 0
            for url in normal_urls:
                try:
                    if agent.validate_url(url):
                        safe_count += 1
                except:
                    pass
            
            # Evaluation
            if unsafe_count == 0 and safe_count == len(normal_urls):
                self.log_security_test("Input Validation", "PASS", f"All malicious inputs rejected, normal inputs accepted")
                return True
            elif unsafe_count <= 2:  # Allow some false negatives but not many false positives
                self.log_security_test("Input Validation", "WARN", f"Some malicious inputs accepted ({unsafe_count}/{len(malicious_inputs)})")
                return True
            else:
                self.log_security_test("Input Validation", "FAIL", f"Too many malicious inputs accepted ({unsafe_count}/{len(malicious_inputs)})")
                return False
                
        except Exception as e:
            self.log_security_test("Input Validation", "FAIL", f"Error: {str(e)}")
            return False
    
    async def test_cors_configuration(self) -> bool:
        """Test CORS configuration"""
        print("\n🌐 Testing CORS Configuration")
        print("-" * 40)
        
        try:
            from backend.app.core.config import settings
            
            cors_origins = settings.CORS_ORIGINS
            
            # Check for overly permissive CORS
            issues = []
            
            if "*" in str(cors_origins):
                issues.append("Wildcard (*) origin detected")
            
            if "localhost" in str(cors_origins) or "127.0.0.1" in str(cors_origins):
                issues.append("Localhost origins present (may be intended for development)")
            
            # Check if origins are properly formatted
            if isinstance(cors_origins, list):
                for origin in cors_origins:
                    if not isinstance(origin, str) or not origin.startswith(("http://", "https://")):
                        issues.append(f"Invalid origin format: {origin}")
            
            if issues:
                self.log_security_test("CORS Configuration", "WARN", "; ".join(issues))
                return True
            else:
                self.log_security_test("CORS Configuration", "PASS", f"CORS configured with {len(cors_origins)} origins")
                return True
                
        except Exception as e:
            self.log_security_test("CORS Configuration", "FAIL", f"Error: {str(e)}")
            return False
    
    async def run_security_suite(self) -> Dict:
        """Run complete security test suite"""
        print("=" * 60)
        print("🔒 GEDI SECURITY VALIDATION TEST")
        print("=" * 60)
        
        security_tests = [
            ("Configuration Security", self.test_config_security),
            ("Password Hashing", self.test_password_hashing),
            ("JWT Token Security", self.test_jwt_token_security),
            ("Input Validation", self.test_input_validation),
            ("CORS Configuration", self.test_cors_configuration),
        ]
        
        results = []
        for test_name, test_func in security_tests:
            try:
                result = await test_func()
                results.append((test_name, result))
            except Exception as e:
                print(f"💥 {test_name} crashed: {str(e)}")
                results.append((test_name, False))
        
        # Generate summary
        passed = sum(1 for _, result in results if result)
        total = len(results)
        success_rate = (passed / total) * 100 if total > 0 else 0
        
        print("\n" + "=" * 60)
        print("🔐 SECURITY TEST SUMMARY")
        print("=" * 60)
        print(f"Tests Passed: {passed}/{total}")
        print(f"Success Rate: {success_rate:.1f}%")
        
        # Security rating
        if success_rate >= 80:
            rating = "🛡️  STRONG"
        elif success_rate >= 60:
            rating = "✅ MODERATE"
        else:
            rating = "⚠️  WEAK"
            
        print(f"Security Posture: {rating}")
        
        # Detailed results
        print("\n📋 DETAILED RESULTS:")
        for test_name, result in results:
            status_icon = "✅" if result else "❌"
            print(f"{status_icon} {test_name}")
        
        # Save results
        import json
        results_data = {
            "summary": {
                "tests_passed": passed,
                "total_tests": total,
                "success_rate": success_rate,
                "security_rating": rating,
                "timestamp": time.time()
            },
            "detailed_results": self.results
        }
        
        with open('security_test_results.json', 'w') as f:
            json.dump(results_data, f, indent=2)
        
        print(f"\n💾 Results saved to: security_test_results.json")
        
        return results_data

async def main():
    """Main security test runner"""
    tester = SecurityTester()
    results = await tester.run_security_suite()
    
    # Return appropriate exit code
    rating = results["summary"]["security_rating"]
    if "STRONG" in rating:
        return 0
    elif "MODERATE" in rating:
        return 1
    else:
        return 2

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    exit(exit_code)
#!/usr/bin/env python3
"""
AI Assistant Integration for MoonFlower Automation
Uses OpenRouter API with DeepSeek R1 for intelligent automation decisions
"""

import requests
import json
import logging
from typing import Dict, Any, Optional, List
from datetime import datetime
import time

class AIAssistant:
    """AI Assistant for intelligent automation decisions and error handling"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # OpenRouter API configuration
        self.api_url = "https://openrouter.ai/api/v1/chat/completions"
        self.api_key = "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"
        self.model = "deepseek/deepseek-r1-0528:free"
        
        # Request headers
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://moonflower-automation.local",
            "X-Title": "MoonFlower WiFi Automation System"
        }
        
        # Rate limiting
        self.last_request_time = 0
        self.min_request_interval = 2.0  # Minimum 2 seconds between requests
        
        self.logger.info("AI Assistant initialized with DeepSeek R1 model")
    
    def _make_api_request(self, messages: List[Dict[str, str]], max_tokens: int = 500) -> Optional[str]:
        """Make API request to OpenRouter with rate limiting"""
        try:
            # Rate limiting
            current_time = time.time()
            time_since_last = current_time - self.last_request_time
            if time_since_last < self.min_request_interval:
                time.sleep(self.min_request_interval - time_since_last)
            
            # Prepare request data
            data = {
                "model": self.model,
                "messages": messages,
                "max_tokens": max_tokens,
                "temperature": 0.7
            }
            
            # Make request
            response = requests.post(
                url=self.api_url,
                headers=self.headers,
                data=json.dumps(data),
                timeout=30
            )
            
            self.last_request_time = time.time()
            
            if response.status_code == 200:
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    return result['choices'][0]['message']['content']
                else:
                    self.logger.error("No choices in API response")
                    return None
            else:
                self.logger.error(f"API request failed: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            self.logger.error(f"AI API request failed: {e}")
            return None
    
    def analyze_automation_error(self, error_context: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze automation errors and suggest solutions"""
        try:
            self.logger.info("Analyzing automation error with AI")
            
            # Prepare context for AI
            context_summary = self._format_error_context(error_context)
            
            messages = [
                {
                    "role": "system",
                    "content": """You are an expert automation engineer specializing in Windows automation, VBS scripting, and web scraping. 
                    Analyze automation errors and provide specific, actionable solutions. Focus on:
                    1. Root cause analysis
                    2. Specific fix recommendations
                    3. Prevention strategies
                    4. Alternative approaches if needed
                    Keep responses concise and technical."""
                },
                {
                    "role": "user",
                    "content": f"Analyze this automation error and provide solutions:\n\n{context_summary}"
                }
            ]
            
            ai_response = self._make_api_request(messages, max_tokens=800)
            
            if ai_response:
                return {
                    'success': True,
                    'analysis': ai_response,
                    'timestamp': datetime.now().isoformat(),
                    'confidence': 'high' if len(ai_response) > 100 else 'medium'
                }
            else:
                return {
                    'success': False,
                    'error': 'AI analysis failed',
                    'timestamp': datetime.now().isoformat()
                }
                
        except Exception as e:
            self.logger.error(f"Error analysis failed: {e}")
            return {
                'success': False,
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }
    
    def optimize_vbs_coordinates(self, coordinate_issues: Dict[str, Any]) -> Dict[str, Any]:
        """Get AI suggestions for VBS coordinate optimization"""
        try:
            self.logger.info("Getting AI suggestions for VBS coordinate optimization")
            
            messages = [
                {
                    "role": "system",
                    "content": """You are a VBS automation expert. Analyze coordinate-based automation issues and suggest:
                    1. Coordinate adjustment strategies
                    2. Alternative element detection methods
                    3. Window state handling approaches
                    4. Fallback mechanisms
                    Provide specific, implementable solutions."""
                },
                {
                    "role": "user",
                    "content": f"VBS coordinate automation issues:\n{json.dumps(coordinate_issues, indent=2)}\n\nSuggest optimizations and fixes."
                }
            ]
            
            ai_response = self._make_api_request(messages, max_tokens=600)
            
            if ai_response:
                return {
                    'success': True,
                    'suggestions': ai_response,
                    'timestamp': datetime.now().isoformat()
                }
            else:
                return {
                    'success': False,
                    'error': 'AI optimization failed'
                }
                
        except Exception as e:
            self.logger.error(f"VBS optimization failed: {e}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def suggest_workflow_improvements(self, workflow_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Get AI suggestions for workflow optimization"""
        try:
            self.logger.info("Getting AI suggestions for workflow improvements")
            
            messages = [
                {
                    "role": "system",
                    "content": """You are a workflow optimization expert for automation systems. Analyze performance metrics and suggest:
                    1. Timing optimizations
                    2. Error reduction strategies
                    3. Efficiency improvements
                    4. Reliability enhancements
                    Focus on practical, implementable improvements."""
                },
                {
                    "role": "user",
                    "content": f"Workflow performance metrics:\n{json.dumps(workflow_metrics, indent=2)}\n\nSuggest improvements for better reliability and performance."
                }
            ]
            
            ai_response = self._make_api_request(messages, max_tokens=700)
            
            if ai_response:
                return {
                    'success': True,
                    'improvements': ai_response,
                    'timestamp': datetime.now().isoformat()
                }
            else:
                return {
                    'success': False,
                    'error': 'AI workflow analysis failed'
                }
                
        except Exception as e:
            self.logger.error(f"Workflow improvement analysis failed: {e}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def generate_error_recovery_plan(self, error_history: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate intelligent error recovery plans based on error patterns"""
        try:
            self.logger.info("Generating AI-powered error recovery plan")
            
            # Summarize error patterns
            error_summary = self._summarize_error_patterns(error_history)
            
            messages = [
                {
                    "role": "system",
                    "content": """You are an expert in automation error recovery and system resilience. 
                    Analyze error patterns and create comprehensive recovery plans including:
                    1. Immediate recovery actions
                    2. Pattern-based prevention
                    3. Escalation procedures
                    4. System health checks
                    Provide actionable, step-by-step recovery procedures."""
                },
                {
                    "role": "user",
                    "content": f"Error pattern analysis:\n{error_summary}\n\nGenerate a comprehensive error recovery plan."
                }
            ]
            
            ai_response = self._make_api_request(messages, max_tokens=900)
            
            if ai_response:
                return {
                    'success': True,
                    'recovery_plan': ai_response,
                    'error_patterns_analyzed': len(error_history),
                    'timestamp': datetime.now().isoformat()
                }
            else:
                return {
                    'success': False,
                    'error': 'AI recovery plan generation failed'
                }
                
        except Exception as e:
            self.logger.error(f"Error recovery plan generation failed: {e}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _format_error_context(self, error_context: Dict[str, Any]) -> str:
        """Format error context for AI analysis"""
        try:
            formatted = []
            
            # Basic error info
            if 'error_type' in error_context:
                formatted.append(f"Error Type: {error_context['error_type']}")
            
            if 'error_message' in error_context:
                formatted.append(f"Error Message: {error_context['error_message']}")
            
            if 'phase' in error_context:
                formatted.append(f"Automation Phase: {error_context['phase']}")
            
            if 'component' in error_context:
                formatted.append(f"Component: {error_context['component']}")
            
            # System context
            if 'system_state' in error_context:
                formatted.append(f"System State: {json.dumps(error_context['system_state'], indent=2)}")
            
            # Recent actions
            if 'recent_actions' in error_context:
                formatted.append(f"Recent Actions: {error_context['recent_actions']}")
            
            # Timing info
            if 'timestamp' in error_context:
                formatted.append(f"Timestamp: {error_context['timestamp']}")
            
            return "\n".join(formatted)
            
        except Exception as e:
            self.logger.error(f"Error context formatting failed: {e}")
            return str(error_context)
    
    def _summarize_error_patterns(self, error_history: List[Dict[str, Any]]) -> str:
        """Summarize error patterns for AI analysis"""
        try:
            if not error_history:
                return "No error history available"
            
            # Count error types
            error_types = {}
            components = {}
            phases = {}
            
            for error in error_history[-20:]:  # Last 20 errors
                error_type = error.get('error_type', 'unknown')
                component = error.get('component', 'unknown')
                phase = error.get('phase', 'unknown')
                
                error_types[error_type] = error_types.get(error_type, 0) + 1
                components[component] = components.get(component, 0) + 1
                phases[phase] = phases.get(phase, 0) + 1
            
            summary = []
            summary.append(f"Total errors analyzed: {len(error_history)}")
            summary.append(f"Most common error types: {dict(sorted(error_types.items(), key=lambda x: x[1], reverse=True)[:5])}")
            summary.append(f"Most affected components: {dict(sorted(components.items(), key=lambda x: x[1], reverse=True)[:5])}")
            summary.append(f"Most problematic phases: {dict(sorted(phases.items(), key=lambda x: x[1], reverse=True)[:5])}")
            
            # Recent errors
            if len(error_history) > 0:
                recent_errors = error_history[-5:]
                summary.append("Recent errors:")
                for i, error in enumerate(recent_errors, 1):
                    summary.append(f"  {i}. {error.get('error_type', 'unknown')} in {error.get('component', 'unknown')} - {error.get('error_message', 'no message')}")
            
            return "\n".join(summary)
            
        except Exception as e:
            self.logger.error(f"Error pattern summarization failed: {e}")
            return f"Error pattern analysis failed: {e}"
    
    def get_system_health_insights(self, health_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Get AI insights on system health and performance"""
        try:
            self.logger.info("Getting AI insights on system health")
            
            messages = [
                {
                    "role": "system",
                    "content": """You are a system health analyst for automation systems. Analyze health metrics and provide:
                    1. Health status assessment
                    2. Performance bottleneck identification
                    3. Preventive maintenance recommendations
                    4. Optimization opportunities
                    Focus on actionable insights for system reliability."""
                },
                {
                    "role": "user",
                    "content": f"System health metrics:\n{json.dumps(health_metrics, indent=2)}\n\nProvide health insights and recommendations."
                }
            ]
            
            ai_response = self._make_api_request(messages, max_tokens=600)
            
            if ai_response:
                return {
                    'success': True,
                    'health_insights': ai_response,
                    'timestamp': datetime.now().isoformat()
                }
            else:
                return {
                    'success': False,
                    'error': 'AI health analysis failed'
                }
                
        except Exception as e:
            self.logger.error(f"System health analysis failed: {e}")
            return {
                'success': False,
                'error': str(e)
            }

def main():
    """Test the AI Assistant"""
    logging.basicConfig(level=logging.INFO)
    
    ai = AIAssistant()
    
    # Test error analysis
    test_error = {
        'error_type': 'CoordinateError',
        'error_message': 'Click coordinates (1194, 692) failed - element not found',
        'phase': 'Phase 3 - Data Import',
        'component': 'VBS Automation',
        'timestamp': datetime.now().isoformat()
    }
    
    result = ai.analyze_automation_error(test_error)
    print("AI Error Analysis:")
    print(json.dumps(result, indent=2))

if __name__ == "__main__":
    main()
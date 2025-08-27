import requests
import json
import time
import logging
from typing import Dict, List, Optional, Any

class EnhancedLLMAssistant:
    def __init__(self):
        self.api_key = "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"
        self.base_url = "https://openrouter.ai/api/v1/chat/completions"
        self.model = "deepseek/deepseek-r1-0528:free"
        self.logger = logging.getLogger(__name__)
        
    def make_request(self, messages: List[Dict], max_retries: int = 3) -> Optional[str]:
        """Make API request with retry logic"""
        for attempt in range(max_retries):
            try:
                response = requests.post(
                    url=self.base_url,
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json",
                        "HTTP-Referer": "moonflower-automation",
                        "X-Title": "MoonFlower Automation System"
                    },
                    data=json.dumps({
                        "model": self.model,
                        "messages": messages,
                        "temperature": 0.1,
                        "max_tokens": 1000
                    }),
                    timeout=30
                )
                
                if response.status_code == 200:
                    result = response.json()
                    return result['choices'][0]['message']['content']
                else:
                    self.logger.error(f"API request failed: {response.status_code} - {response.text}")
                    
            except Exception as e:
                self.logger.error(f"Request attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)  # Exponential backoff
                    
        return None
    
    def analyze_error_context(self, error_details: Dict) -> str:
        """Analyze error context and provide actionable insights"""
        messages = [
            {
                "role": "system",
                "content": "You are an automation error analyst. Provide concise, actionable solutions for VBS automation errors."
            },
            {
                "role": "user", 
                "content": f"Analyze this automation error and suggest recovery: {json.dumps(error_details)}"
            }
        ]
        
        response = self.make_request(messages)
        return response or "Unable to analyze error - using fallback recovery"
    
    def optimize_timing_strategy(self, phase: str, historical_data: Dict) -> Dict:
        """Get LLM-optimized timing strategy for phases"""
        messages = [
            {
                "role": "system",
                "content": "You are a timing optimization expert for VBS automation. Return JSON with wait_times, retry_counts, and timeout_values."
            },
            {
                "role": "user",
                "content": f"Optimize timing for {phase} based on: {json.dumps(historical_data)}"
            }
        ]
        
        response = self.make_request(messages)
        try:
            if response:
                return json.loads(response)
        except:
            pass
        
        # Fallback timing strategy
        return {
            "wait_times": {"element_load": 2, "action_delay": 1, "page_transition": 3},
            "retry_counts": {"element_find": 3, "action_perform": 2},
            "timeout_values": {"element_presence": 10, "page_load": 30}
        }
    
    def generate_dynamic_selectors(self, page_context: str, target_element: str) -> List[str]:
        """Generate alternative selectors for robust element detection"""
        messages = [
            {
                "role": "system", 
                "content": "Generate alternative CSS/XPath selectors for UI automation. Return as JSON array."
            },
            {
                "role": "user",
                "content": f"Page context: {page_context}. Target: {target_element}. Provide 3-5 alternative selectors."
            }
        ]
        
        response = self.make_request(messages)
        try:
            if response:
                return json.loads(response)
        except:
            pass
        
        return [f"[data-testid='{target_element}']", f".{target_element}", f"#{target_element}"]
    
    def validate_phase_completion(self, phase: str, captured_state: Dict) -> bool:
        """Validate if phase completed successfully using LLM analysis"""
        messages = [
            {
                "role": "system",
                "content": "Validate VBS automation phase completion. Return 'SUCCESS' or 'FAILURE' with brief reason."
            },
            {
                "role": "user",
                "content": f"Phase: {phase}. State: {json.dumps(captured_state)}"
            }
        ]
        
        response = self.make_request(messages)
        return response and "SUCCESS" in response.upper()
    
    def adaptive_recovery_strategy(self, failure_context: Dict) -> Dict:
        """Get adaptive recovery strategy from LLM"""
        messages = [
            {
                "role": "system",
                "content": "Provide VBS automation recovery strategy as JSON with steps, timeouts, and alternative_approaches."
            },
            {
                "role": "user",
                "content": f"Failure context: {json.dumps(failure_context)}"
            }
        ]
        
        response = self.make_request(messages)
        try:
            if response:
                return json.loads(response)
        except:
            pass
        
        return {
            "steps": ["restart_application", "retry_login", "resume_from_checkpoint"],
            "timeouts": {"restart": 30, "retry": 15},
            "alternative_approaches": ["keyboard_navigation", "coordinate_fallback"]
        }
#!/usr/bin/env python3
"""
Enhanced VBS Audio Detection Module
Optimized for detecting clicks, pops, and short audio cues
"""

import time
import logging
import threading
import os
from datetime import datetime
from typing import Optional, Callable, Dict, Any
import traceback

# Try to import audio libraries
try:
    import pyaudio
    import numpy as np
    from scipy import signal
    AUDIO_LIBS_AVAILABLE = True
    SCIPY_AVAILABLE = True
except ImportError as e:
    try:
        import pyaudio
        import numpy as np
        AUDIO_LIBS_AVAILABLE = True
        SCIPY_AVAILABLE = False
    except ImportError:
        AUDIO_LIBS_AVAILABLE = False
        SCIPY_AVAILABLE = False

class EnhancedVBSAudioDetector:
    """Enhanced audio detection system optimized for clicks, pops, and short sounds"""
    
    def __init__(self, vbs_window_handle: Optional[int] = None):
        """Initialize Enhanced VBS Audio Detector"""
        self.logger = self._setup_logging()
        self.vbs_window_handle = vbs_window_handle
        
        # Enhanced audio configuration for POPUP-SPECIFIC detection
        self.config = {
            "sample_rate": 44100,
            "channels": 1,
            "chunk_size": 512,  # Smaller chunk for better temporal resolution
            "format": pyaudio.paInt16 if AUDIO_LIBS_AVAILABLE else None,
            
            # POPUP-SPECIFIC Detection Parameters (more restrictive than generic clicks)
            "rms_threshold": 0.05,          # Higher threshold to ignore typing/clicking
            "peak_threshold": 0.4,          # Higher peak threshold for popup sounds
            "transient_threshold": 0.2,     # Higher threshold for sudden changes
            "min_duration": 0.1,            # 100ms minimum (popup sounds are longer than clicks)
            "max_duration": 1.5,            # Maximum popup sound duration (shorter than before)
            "silence_reset": 1.0,           # Longer reset time to avoid rapid false triggers
            
            # Advanced Detection Features for POPUP sounds
            "use_peak_detection": True,     # Enable peak-based detection
            "use_transient_detection": True, # Enable transient detection
            "use_spectral_detection": True, # Enable frequency analysis for popup sounds
            "use_sustained_detection": True, # New: detect sustained sounds (popups vs clicks)
            
            # Noise Filtering (more aggressive)
            "noise_gate_threshold": 0.01,   # Higher noise gate to ignore background sounds
            "adaptive_threshold": True,     # Adapt threshold based on background noise
            "frequency_filter": True,       # New: filter by frequency content
            "min_popup_frequency": 800,     # Minimum frequency for popup sounds (Hz)
            "max_popup_frequency": 8000,    # Maximum frequency for popup sounds (Hz)
        }
        
        # Detection state
        self.is_detecting = False
        self.success_detected = False
        self.detection_timestamp = None
        self.success_callback = None
        self.detection_count = 0
        
        # Background noise estimation
        self.background_noise_level = 0.005
        self.noise_samples = []
        self.calibration_complete = False
        
        # Audio system
        self.audio_system = None
        self.stream = None
        self.detection_thread = None
        
        # Fallback detection
        self.fallback_mode = not AUDIO_LIBS_AVAILABLE
        
        self.logger.info(f"🔊 Enhanced VBS Audio Detector initialized")
        self.logger.info(f"   Audio libs: {'Available' if AUDIO_LIBS_AVAILABLE else 'Fallback mode'}")
        self.logger.info(f"   SciPy: {'Available' if SCIPY_AVAILABLE else 'Not available'}")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup enhanced logging"""
        logger = logging.getLogger("EnhancedVBSAudioDetector")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            try:
                log_file = "EHC_Logs/enhanced_vbs_audio_detector.log"
                os.makedirs(os.path.dirname(log_file), exist_ok=True)
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
            except Exception:
                pass
        
        return logger
    
    def calibrate_background_noise(self, duration: float = 3.0) -> float:
        """Calibrate background noise level"""
        try:
            if not AUDIO_LIBS_AVAILABLE:
                self.logger.warning("⚠️ Audio libs not available, skipping calibration")
                return self.background_noise_level
            
            self.logger.info(f"🎯 Calibrating background noise level ({duration}s)...")
            
            # Temporarily open audio stream for calibration
            stream = self.audio_system.open(
                format=self.config["format"],
                channels=self.config["channels"],
                rate=self.config["sample_rate"],
                input=True,
                frames_per_buffer=self.config["chunk_size"]
            )
            
            noise_samples = []
            start_time = time.time()
            
            while (time.time() - start_time) < duration:
                try:
                    data = stream.read(self.config["chunk_size"], exception_on_overflow=False)
                    audio_data = np.frombuffer(data, dtype=np.int16)
                    rms = np.sqrt(np.mean(audio_data**2)) / 32768.0
                    noise_samples.append(rms)
                except Exception as e:
                    self.logger.warning(f"⚠️ Calibration read error: {e}")
                
                time.sleep(0.01)
            
            stream.close()
            
            if noise_samples:
                # Use 95th percentile as background noise level to avoid outliers
                self.background_noise_level = np.percentile(noise_samples, 95)
                
                # Adjust thresholds based on background noise (MORE RESTRICTIVE for popups)
                if self.config["adaptive_threshold"]:
                    self.config["rms_threshold"] = max(0.05, self.background_noise_level * 5)  # Higher multiplier
                    self.config["noise_gate_threshold"] = max(0.01, self.background_noise_level * 3)  # Higher gate
                
                self.calibration_complete = True
                self.logger.info(f"✅ Background noise calibrated: {self.background_noise_level:.4f}")
                self.logger.info(f"   Adjusted RMS threshold: {self.config['rms_threshold']:.4f}")
                
                return self.background_noise_level
            else:
                self.logger.warning("⚠️ No noise samples collected during calibration")
                return self.background_noise_level
                
        except Exception as e:
            self.logger.error(f"❌ Background noise calibration failed: {e}")
            return self.background_noise_level
    
    def initialize_audio_system(self) -> Dict[str, Any]:
        """Initialize enhanced audio detection system"""
        try:
            if not AUDIO_LIBS_AVAILABLE:
                self.logger.warning("⚠️ Audio libraries not available, using fallback detection")
                return {
                    "success": True,
                    "method": "fallback",
                    "message": "Using fallback detection method"
                }
            
            # Initialize PyAudio
            self.audio_system = pyaudio.PyAudio()
            
            # Get default input device info
            default_device = self.audio_system.get_default_input_device_info()
            self.logger.info(f"🎤 Default audio input: {default_device['name']}")
            
            # Test audio stream creation
            test_stream = self.audio_system.open(
                format=self.config["format"],
                channels=self.config["channels"],
                rate=self.config["sample_rate"],
                input=True,
                frames_per_buffer=self.config["chunk_size"]
            )
            test_stream.close()
            
            # Calibrate background noise
            self.calibrate_background_noise()
            
            self.logger.info("✅ Enhanced audio system initialized successfully")
            return {
                "success": True,
                "method": "audio",
                "device": default_device['name'],
                "sample_rate": self.config["sample_rate"],
                "background_noise": self.background_noise_level,
                "scipy_available": SCIPY_AVAILABLE
            }
            
        except Exception as e:
            self.logger.error(f"❌ Audio system initialization failed: {e}")
            self.fallback_mode = True
            return {
                "success": True,
                "method": "fallback",
                "message": f"Audio init failed, using fallback: {str(e)}"
            }
    
    def _detect_click_pop(self, audio_data: np.ndarray) -> tuple[bool, float, str]:
        """Enhanced POPUP sound detection using multiple methods with higher selectivity"""
        try:
            # Normalize audio data
            normalized_audio = audio_data.astype(np.float32) / 32768.0
            
            # Method 1: RMS Detection (for volume-based detection) - MORE RESTRICTIVE
            rms = np.sqrt(np.mean(normalized_audio**2))
            rms_detected = rms > self.config["rms_threshold"]
            
            # Method 2: Peak Detection (for sharp transients) - MORE RESTRICTIVE
            peak_detected = False
            peak_value = 0.0
            if self.config["use_peak_detection"]:
                peak_value = np.max(np.abs(normalized_audio))
                peak_detected = peak_value > self.config["peak_threshold"]
            
            # Method 3: Transient Detection (for sudden changes) - MORE RESTRICTIVE
            transient_detected = False
            transient_strength = 0.0
            if self.config["use_transient_detection"] and len(normalized_audio) > 1:
                # Calculate energy difference between consecutive samples
                energy_diff = np.diff(normalized_audio**2)
                transient_strength = np.max(np.abs(energy_diff))
                transient_detected = transient_strength > self.config["transient_threshold"]
            
            # Method 4: Enhanced Spectral Detection for POPUP sounds
            spectral_detected = False
            spectral_score = 0.0
            if self.config["use_spectral_detection"] and SCIPY_AVAILABLE:
                try:
                    # Analyze frequency content for popup-specific characteristics
                    freqs, psd = signal.welch(normalized_audio, self.config["sample_rate"], nperseg=min(256, len(normalized_audio)))
                    
                    # Focus on popup frequency range (800Hz - 8kHz)
                    popup_freq_mask = (freqs >= self.config["min_popup_frequency"]) & (freqs <= self.config["max_popup_frequency"])
                    popup_energy = np.sum(psd[popup_freq_mask])
                    
                    # Look for sustained energy in popup frequency range
                    total_energy = np.sum(psd)
                    if total_energy > 0:
                        popup_ratio = popup_energy / total_energy
                        spectral_score = popup_ratio
                        spectral_detected = popup_ratio > 0.3  # At least 30% of energy in popup range
                except Exception:
                    pass
            
            # Method 5: Sustained Detection (NEW) - for detecting longer popup sounds vs brief clicks
            sustained_detected = False
            sustained_score = 0.0
            if self.config["use_sustained_detection"]:
                # Calculate energy variance to detect sustained vs transient sounds
                if len(normalized_audio) > 10:
                    energy_variance = np.var(normalized_audio**2)
                    # Lower variance indicates more sustained sound (like a popup beep)
                    sustained_score = 1.0 / (1.0 + energy_variance * 1000)  # Inverse relationship
                    sustained_detected = sustained_score > 0.5 and rms > 0.02
            
            # Combine detection methods with STRICTER requirements
            detection_score = max(rms, peak_value, transient_strength, spectral_score, sustained_score)
            
            # STRICTER OVERALL DETECTION - require multiple criteria OR very high single score
            multiple_detections = sum([rms_detected, peak_detected, transient_detected, spectral_detected, sustained_detected])
            high_confidence = detection_score > 0.6
            
            overall_detected = (multiple_detections >= 2) or high_confidence
            
            # Apply STRICTER noise gate
            if detection_score < self.config["noise_gate_threshold"]:
                overall_detected = False
            
            # Additional filtering: reject very brief or very long sounds
            # (This will be applied later in the duration check)
            
            # Determine detection method used
            method = []
            if rms_detected: method.append("RMS")
            if peak_detected: method.append("Peak")
            if transient_detected: method.append("Transient")
            if spectral_detected: method.append("Spectral")
            if sustained_detected: method.append("Sustained")
            method_str = "+".join(method) if method else "None"
            
            return overall_detected, detection_score, method_str
            
        except Exception as e:
            self.logger.warning(f"⚠️ Popup detection error: {e}")
            return False, 0.0, "Error"
    
    def _audio_detection_worker(self, timeout: Optional[float] = None):
        """Enhanced audio detection worker thread"""
        try:
            self.logger.info("🎧 Enhanced audio detection worker started")
            
            # Open audio stream with smaller buffer for better responsiveness
            self.stream = self.audio_system.open(
                format=self.config["format"],
                channels=self.config["channels"],
                rate=self.config["sample_rate"],
                input=True,
                frames_per_buffer=self.config["chunk_size"]
            )
            
            start_time = time.time()
            sound_start_time = None
            sound_chunks = []
            detection_active = False
            
            while self.is_detecting:
                # Check timeout within the worker thread
                if timeout and (time.time() - start_time) > timeout:
                    self.logger.info("⏰ Enhanced audio detection timeout reached in worker")
                    break
                
                # Exit immediately if success already detected by another thread
                if self.success_detected:
                    self.logger.info("🔔 Success already detected, exiting worker")
                    break
                
                try:
                    # Read audio data
                    data = self.stream.read(self.config["chunk_size"], exception_on_overflow=False)
                    audio_data = np.frombuffer(data, dtype=np.int16)
                    
                    # Detect click/pop
                    detected, score, method = self._detect_click_pop(audio_data)
                    
                    if detected:
                        if not detection_active:
                            # Start of new detection
                            sound_start_time = time.time()
                            detection_active = True
                            sound_chunks = [audio_data]
                            self.logger.info(f"🔍 Sound detected (Method: {method}, Score: {score:.4f})")
                        else:
                            # Continue existing detection
                            sound_chunks.append(audio_data)
                    else:
                        if detection_active:
                            # Check if we have a valid detection duration
                            sound_duration = time.time() - sound_start_time
                            
                            if (self.config["min_duration"] <= sound_duration <= self.config["max_duration"]):
                                # Valid popup sound detected!
                                self.detection_count += 1
                                self.logger.info(f"🔔 POPUP SOUND #{self.detection_count} detected! Duration: {sound_duration:.3f}s")
                                self.logger.info(f"   Detection criteria met: {method}")
                                self.logger.info(f"   Sound duration: {sound_duration:.3f}s (range: {self.config['min_duration']:.3f}s - {self.config['max_duration']:.3f}s)")
                                self._trigger_success_detection()
                                # Exit immediately after successful detection
                                break
                            else:
                                # Invalid duration, reset
                                if sound_duration < self.config["min_duration"]:
                                    self.logger.debug(f"🔇 Sound too short: {sound_duration:.3f}s")
                                else:
                                    self.logger.debug(f"🔇 Sound too long: {sound_duration:.3f}s")
                            
                            detection_active = False
                            sound_chunks = []
                    
                    # Minimal delay for maximum responsiveness
                    time.sleep(0.001)  # 1ms delay for maximum responsiveness
                    
                except Exception as e:
                    self.logger.warning(f"⚠️ Audio read error: {e}")
                    time.sleep(0.1)
            
            # Ensure detection is stopped when worker exits
            self.is_detecting = False
            
        except Exception as e:
            self.logger.error(f"❌ Enhanced audio detection worker failed: {e}")
            self.is_detecting = False
        finally:
            if self.stream:
                try:
                    self.stream.stop_stream()
                    self.stream.close()
                except:
                    pass
                self.stream = None
            self.logger.info("🎧 Enhanced audio detection worker finished")
    
    def start_detection(self, success_callback: Optional[Callable] = None, timeout: Optional[float] = None) -> bool:
        """Start enhanced audio detection"""
        try:
            if self.is_detecting:
                self.logger.warning("⚠️ Detection already in progress")
                return False
            
            self.success_callback = success_callback
            self.success_detected = False
            self.detection_timestamp = None
            self.detection_count = 0
            self.is_detecting = True
            
            if self.fallback_mode:
                self.logger.info("🔄 Starting fallback detection mode")
                self.detection_thread = threading.Thread(
                    target=self._fallback_detection_worker,
                    args=(timeout,),
                    daemon=True
                )
            else:
                self.logger.info("🔄 Starting enhanced audio detection mode")
                self.detection_thread = threading.Thread(
                    target=self._audio_detection_worker,
                    args=(timeout,),
                    daemon=True
                )
            
            self.detection_thread.start()
            self.logger.info("✅ Enhanced audio detection started")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Failed to start detection: {e}")
            self.is_detecting = False
            return False
    
    def stop_detection(self) -> bool:
        """Stop audio detection"""
        try:
            if not self.is_detecting:
                return True
            
            self.is_detecting = False
            
            # Close audio stream if active
            if self.stream:
                self.stream.stop_stream()
                self.stream.close()
                self.stream = None
            
            # Wait for detection thread to finish
            if self.detection_thread and self.detection_thread.is_alive():
                self.detection_thread.join(timeout=2.0)
            
            self.logger.info("✅ Enhanced audio detection stopped")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Failed to stop detection: {e}")
            return False
    
    def _fallback_detection_worker(self, timeout: Optional[float] = None):
        """Fallback detection worker using time-based simulation"""
        try:
            self.logger.info("🔄 Fallback detection worker started")
            
            start_time = time.time()
            detection_delay = 30.0  # Shorter delay for click detection
            
            while self.is_detecting:
                elapsed = time.time() - start_time
                
                if timeout and elapsed > timeout:
                    self.logger.info("⏰ Fallback detection timeout reached")
                    break
                
                if elapsed >= detection_delay:
                    self.logger.info("🔔 Fallback click detection triggered")
                    self._trigger_success_detection()
                    break
                
                if int(elapsed) % 5 == 0 and int(elapsed) > 0:
                    remaining = max(0, detection_delay - elapsed)
                    self.logger.info(f"⏱️ Fallback detection: {remaining:.0f}s remaining")
                
                time.sleep(1.0)
                
        except Exception as e:
            self.logger.error(f"❌ Fallback detection worker failed: {e}")
        finally:
            self.logger.info("🔄 Fallback detection worker finished")
    
    def _trigger_success_detection(self):
        """Trigger success detection event"""
        try:
            self.success_detected = True
            self.detection_timestamp = datetime.now()
            self.is_detecting = False
            
            self.logger.info(f"🎉 Success detection triggered at {self.detection_timestamp}")
            
            # Call success callback if provided
            if self.success_callback:
                try:
                    self.success_callback()
                except Exception as e:
                    self.logger.error(f"❌ Success callback failed: {e}")
            
        except Exception as e:
            self.logger.error(f"❌ Failed to trigger success detection: {e}")
    
    def wait_for_click_sound(self, timeout: float = 300.0) -> bool:
        """Wait for click/pop sound with timeout"""
        try:
            self.logger.info(f"⏳ Waiting for click/pop sound (timeout: {timeout}s)")
            
            if not self.start_detection(timeout=timeout):
                return False
            
            start_time = time.time()
            while self.is_detecting and (time.time() - start_time) < timeout:
                if self.success_detected:
                    self.logger.info("✅ Click/pop sound detected!")
                    return True
                time.sleep(0.1)  # More frequent checking
            
            # Timeout reached
            self.stop_detection()
            self.logger.warning("⏰ Timeout waiting for click/pop sound")
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Wait for click/pop sound failed: {e}")
            return False
    
    def wait_for_success_sound(self, timeout: float = 300.0) -> bool:
        """Wait for success popup sound (alias for click/pop detection optimized for popups)"""
        try:
            self.logger.info(f"🔔 Waiting for success popup sound (timeout: {timeout}s)")
            
            # Reset detection state for fresh popup detection
            self.success_detected = False
            self.detection_timestamp = None
            
            if not self.start_detection(timeout=timeout):
                return False
            
            start_time = time.time()
            check_interval = 0.1  # More frequent checking for better responsiveness
            last_log_time = 0
            
            while (time.time() - start_time) < timeout:
                elapsed = time.time() - start_time
                
                # Check if success was detected
                if self.success_detected:
                    self.stop_detection()  # Stop detection immediately
                    self.logger.info(f"🔔 SUCCESS: Popup sound detected after {elapsed:.1f}s!")
                    return True
                
                # Check if detection worker has stopped (could indicate success or failure)
                if not self.is_detecting:
                    # Give a moment for success_detected to be set
                    time.sleep(0.1)
                    if self.success_detected:
                        self.logger.info(f"🔔 SUCCESS: Popup sound detected after {elapsed:.1f}s!")
                        return True
                    else:
                        # Worker stopped without success - restart if still within timeout
                        remaining_time = timeout - elapsed
                        if remaining_time > 10:  # Only restart if significant time left
                            self.logger.info(f"🔄 Restarting audio detection, {remaining_time:.1f}s remaining")
                            if not self.start_detection(timeout=remaining_time):
                                break
                        else:
                            break
                
                # Progress logging every 30 seconds
                if elapsed - last_log_time >= 30:
                    remaining = timeout - elapsed
                    self.logger.info(f"🔊 Listening for popup sound: {elapsed:.1f}s elapsed, {remaining:.1f}s remaining")
                    last_log_time = elapsed
                
                time.sleep(check_interval)
            
            # Timeout reached or detection failed
            self.stop_detection()
            self.logger.warning(f"⏰ Timeout waiting for success popup sound after {timeout:.1f}s")
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Wait for success popup sound failed: {e}")
            self.stop_detection()
            return False
    
    def get_detection_status(self) -> Dict[str, Any]:
        """Get current detection status"""
        return {
            "is_detecting": self.is_detecting,
            "success_detected": self.success_detected,
            "detection_count": self.detection_count,
            "detection_timestamp": self.detection_timestamp.isoformat() if self.detection_timestamp else None,
            "fallback_mode": self.fallback_mode,
            "audio_available": AUDIO_LIBS_AVAILABLE,
            "scipy_available": SCIPY_AVAILABLE,
            "background_noise_level": self.background_noise_level,
            "calibration_complete": self.calibration_complete,
            "config": self.config
        }
    
    def cleanup(self):
        """Clean up audio resources"""
        try:
            self.stop_detection()
            
            if self.audio_system:
                self.audio_system.terminate()
                self.audio_system = None
            
            self.logger.info("🧹 Enhanced audio detector cleanup completed")
            
        except Exception as e:
            self.logger.error(f"❌ Enhanced audio detector cleanup failed: {e}")

# Convenience functions
def create_enhanced_audio_detector(vbs_window_handle: Optional[int] = None) -> EnhancedVBSAudioDetector:
    """Create and initialize enhanced audio detector"""
    detector = EnhancedVBSAudioDetector(vbs_window_handle)
    detector.initialize_audio_system()
    return detector

def wait_for_vbs_click_sound(timeout: float = 300.0, vbs_window_handle: Optional[int] = None) -> bool:
    """Wait for VBS click/pop sound with timeout"""
    detector = create_enhanced_audio_detector(vbs_window_handle)
    try:
        return detector.wait_for_click_sound(timeout)
    finally:
        detector.cleanup()

if __name__ == "__main__":
    # Test the enhanced audio detector
    print("🧪 Testing Enhanced VBS Audio Detector")
    print("=" * 50)
    
    detector = EnhancedVBSAudioDetector()
    init_result = detector.initialize_audio_system()
    
    print(f"Initialization: {init_result}")
    print(f"Background noise level: {detector.background_noise_level:.4f}")
    
    # Test detection for 15 seconds
    print("Testing click/pop detection for 15 seconds...")
    print("Make some clicking sounds to test detection!")
    
    result = detector.wait_for_click_sound(timeout=15.0)
    
    print(f"Detection result: {result}")
    print(f"Status: {detector.get_detection_status()}")
    
    detector.cleanup()
    print("Test completed")
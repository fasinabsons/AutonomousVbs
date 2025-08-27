#!/usr/bin/env python3
"""
Audio Detection Service for VBS Automation
Detects click sounds and completion audio cues for Phase 3 upload monitoring
Essential for detecting import success and update completion as specified in vbsupdate.txt
"""

import time
import logging
import threading
import numpy as np
from datetime import datetime
from dataclasses import dataclass
from typing import Optional
import os

try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False
    print("⚠️ PyAudio not available. Audio detection will be disabled.")

@dataclass
class AudioDetectionResult:
    """Result of audio detection operation"""
    detected: bool
    detection_type: str  # "click" or "completion"
    detection_time: Optional[float] = None
    timeout: bool = False
    error: Optional[str] = None

class AudioDetector:
    """
    Audio detection service for monitoring click sounds and completion audio cues
    Essential for Phase 3 upload completion detection as specified in vbsupdate.txt
    """
    
    def __init__(self, config=None):
        self.config = config or {}
        self.logger = self._setup_logging()
        
        # Audio configuration
        self.sample_rate = self.config.get('sample_rate', 44100)
        self.chunk_size = self.config.get('chunk_size', 1024)
        self.channels = 1
        
        # Detection state
        self.is_monitoring = False
        self.click_sound_detected = False
        self.completion_sound_detected = False
        self.audio_buffer = []
        
        # Audio thresholds and patterns
        self.click_sound_threshold = self.config.get('click_sound_threshold', 0.3)
        self.completion_sound_threshold = self.config.get('completion_sound_threshold', 0.2)
        self.background_noise_level = self.config.get('background_noise_level', 0.1)
        
        # PyAudio objects
        self.audio = None
        self.stream = None
        
        # Check PyAudio availability
        if not PYAUDIO_AVAILABLE:
            self.logger.warning("⚠️ PyAudio not available. Audio detection disabled.")
            return
        
        self.logger.info("🔊 Audio detector initialized")
        self.logger.info(f"   Sample rate: {self.sample_rate} Hz")
        self.logger.info(f"   Chunk size: {self.chunk_size}")
        self.logger.info(f"   Click threshold: {self.click_sound_threshold}")
        self.logger.info(f"   Completion threshold: {self.completion_sound_threshold}")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for audio detection"""
        logger = logging.getLogger("AudioDetector")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                log_dir = "EHC_Logs"
                os.makedirs(log_dir, exist_ok=True)
                file_handler = logging.FileHandler(f"{log_dir}/audio_detection.log", encoding='utf-8')
                file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
            except Exception as e:
                print(f"Warning: Could not create audio log file: {e}")
        
        return logger
    
    def is_available(self) -> bool:
        """Check if audio detection is available"""
        return PYAUDIO_AVAILABLE
    
    def calibrate_background_noise(self, duration_seconds=5) -> float:
        """
        Calibrate background noise level for better detection accuracy
        Should be run before starting automation
        """
        if not self.is_available():
            self.logger.warning("Audio detection not available for calibration")
            return 0.1
        
        self.logger.info(f"🔧 Calibrating background noise level for {duration_seconds} seconds...")
        self.logger.info("   Please ensure VBS application is running but idle")
        
        noise_levels = []
        
        try:
            # Initialize audio for calibration
            self.audio = pyaudio.PyAudio()
            stream = self.audio.open(
                format=pyaudio.paInt16,
                channels=self.channels,
                rate=self.sample_rate,
                input=True,
                frames_per_buffer=self.chunk_size
            )
            
            start_time = time.time()
            while time.time() - start_time < duration_seconds:
                try:
                    data = stream.read(self.chunk_size, exception_on_overflow=False)
                    audio_data = np.frombuffer(data, dtype=np.int16)
                    noise_level = np.sqrt(np.mean(audio_data**2)) / 32768.0
                    noise_levels.append(noise_level)
                    time.sleep(0.1)
                except Exception as e:
                    self.logger.warning(f"Audio read error during calibration: {e}")
                    continue
            
            stream.stop_stream()
            stream.close()
            self.audio.terminate()
            
            if noise_levels:
                self.background_noise_level = np.mean(noise_levels) * 1.5  # Add margin
                self.logger.info(f"✅ Background noise level calibrated: {self.background_noise_level:.4f}")
            else:
                self.logger.warning("⚠️ No audio data collected during calibration")
                self.background_noise_level = 0.1
            
        except Exception as e:
            self.logger.error(f"❌ Background noise calibration failed: {e}")
            self.background_noise_level = 0.1
        
        return self.background_noise_level
    
    def start_monitoring(self, duration_seconds=None) -> bool:
        """
        Start audio monitoring for click sounds and completion cues
        Used during Phase 3 update process monitoring
        """
        if not self.is_available():
            self.logger.warning("Audio detection not available")
            return False
        
        try:
            self.audio = pyaudio.PyAudio()
            self.stream = self.audio.open(
                format=pyaudio.paInt16,
                channels=self.channels,
                rate=self.sample_rate,
                input=True,
                frames_per_buffer=self.chunk_size,
                stream_callback=self._audio_callback
            )
            
            self.is_monitoring = True
            self.click_sound_detected = False
            self.completion_sound_detected = False
            self.audio_buffer = []
            
            self.stream.start_stream()
            self.logger.info("🔊 Audio monitoring started")
            
            if duration_seconds:
                # Monitor for specific duration
                time.sleep(duration_seconds)
                self.stop_monitoring()
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Failed to start audio monitoring: {e}")
            return False
    
    def stop_monitoring(self):
        """Stop audio monitoring and cleanup resources"""
        try:
            self.is_monitoring = False
            
            if hasattr(self, 'stream') and self.stream:
                self.stream.stop_stream()
                self.stream.close()
                self.stream = None
            
            if hasattr(self, 'audio') and self.audio:
                self.audio.terminate()
                self.audio = None
            
            self.logger.info("🔇 Audio monitoring stopped")
            
        except Exception as e:
            self.logger.error(f"❌ Error stopping audio monitoring: {e}")
    
    def _audio_callback(self, in_data, frame_count, time_info, status):
        """
        Audio callback function for real-time audio processing
        Detects click sounds and completion audio cues
        """
        if not self.is_monitoring:
            return (in_data, pyaudio.paContinue)
        
        try:
            # Convert audio data to numpy array
            audio_data = np.frombuffer(in_data, dtype=np.int16)
            
            # Calculate audio level (RMS)
            audio_level = np.sqrt(np.mean(audio_data**2)) / 32768.0
            
            # Detect click sound (short, sharp audio spike)
            if self._is_click_sound(audio_data, audio_level):
                self.click_sound_detected = True
                self.logger.info("🔔 Click sound detected!")
            
            # Detect completion sound (longer duration audio pattern)
            if self._is_completion_sound(audio_data, audio_level):
                self.completion_sound_detected = True
                self.logger.info("🎉 Completion sound detected!")
            
            # Store audio buffer for analysis (keep last 5 seconds)
            self.audio_buffer.append(audio_data)
            max_buffer_size = int(self.sample_rate * 5 / self.chunk_size)
            if len(self.audio_buffer) > max_buffer_size:
                self.audio_buffer.pop(0)
            
            return (in_data, pyaudio.paContinue)
            
        except Exception as e:
            self.logger.error(f"❌ Audio callback error: {e}")
            return (in_data, pyaudio.paContinue)
    
    def _is_click_sound(self, audio_data, audio_level) -> bool:
        """
        Detect click sound pattern (short, sharp audio spike)
        Used for detecting import success popup and update completion
        """
        if audio_level < self.click_sound_threshold:
            return False
        
        try:
            # Click sounds typically have higher frequency content
            fft = np.fft.fft(audio_data)
            frequencies = np.fft.fftfreq(len(fft), 1/self.sample_rate)
            
            # Look for higher frequency content typical of click sounds (1kHz - 8kHz)
            high_freq_mask = (frequencies > 1000) & (frequencies < 8000)
            high_freq_power = np.sum(np.abs(fft[high_freq_mask]))
            total_power = np.sum(np.abs(fft))
            
            high_freq_ratio = high_freq_power / total_power if total_power > 0 else 0
            
            # Click sounds have high frequency content and are above background noise
            is_click = (high_freq_ratio > 0.3 and 
                       audio_level > self.background_noise_level * 3)
            
            if is_click:
                self.logger.debug(f"Click detected: level={audio_level:.3f}, freq_ratio={high_freq_ratio:.3f}")
            
            return is_click
            
        except Exception as e:
            self.logger.warning(f"Click detection error: {e}")
            return False
    
    def _is_completion_sound(self, audio_data, audio_level) -> bool:
        """
        Detect completion sound pattern (longer duration audio cue)
        Used for detecting update process completion
        """
        if audio_level < self.completion_sound_threshold:
            return False
        
        try:
            # Check if this is part of a sustained audio pattern
            if len(self.audio_buffer) >= 10:  # At least ~0.25 seconds of history
                recent_levels = []
                for recent_audio in self.audio_buffer[-10:]:
                    recent_level = np.sqrt(np.mean(recent_audio**2)) / 32768.0
                    recent_levels.append(recent_level)
                
                avg_recent_level = np.mean(recent_levels)
                
                # Sustained audio level indicates completion sound
                is_completion = avg_recent_level > self.background_noise_level * 2
                
                if is_completion:
                    self.logger.debug(f"Completion detected: avg_level={avg_recent_level:.3f}")
                
                return is_completion
            
            return False
            
        except Exception as e:
            self.logger.warning(f"Completion detection error: {e}")
            return False
    
    def wait_for_click_sound(self, timeout_seconds=30) -> AudioDetectionResult:
        """
        Wait for click sound detection with timeout
        Used in Phase 3 after import button click - WHEN WE HEAR CLICK+POP SOUND only we press enter
        """
        if not self.is_available():
            return AudioDetectionResult(
                detected=False,
                detection_type="click",
                error="Audio detection not available"
            )
        
        self.logger.info(f"🔊 Waiting for click sound (timeout: {timeout_seconds}s)...")
        
        if not self.start_monitoring():
            return AudioDetectionResult(
                detected=False,
                detection_type="click",
                error="Failed to start audio monitoring"
            )
        
        start_time = time.time()
        while time.time() - start_time < timeout_seconds:
            if self.click_sound_detected:
                detection_time = time.time() - start_time
                self.stop_monitoring()
                self.logger.info(f"✅ Click sound detected after {detection_time:.1f} seconds")
                return AudioDetectionResult(
                    detected=True,
                    detection_type="click",
                    detection_time=detection_time
                )
            time.sleep(0.1)
        
        self.stop_monitoring()
        self.logger.warning(f"⏰ Click sound timeout after {timeout_seconds} seconds")
        return AudioDetectionResult(
            detected=False,
            detection_type="click",
            timeout=True
        )
    
    def wait_for_completion_sound(self, timeout_seconds=7200) -> AudioDetectionResult:
        """
        Wait for completion sound detection with timeout
        Used in Phase 3 during update process monitoring - click+popup sound (20 min-2hours wait)
        """
        if not self.is_available():
            return AudioDetectionResult(
                detected=False,
                detection_type="completion",
                error="Audio detection not available"
            )
        
        self.logger.info(f"🔊 Waiting for completion sound (timeout: {timeout_seconds/3600:.1f} hours)...")
        
        if not self.start_monitoring():
            return AudioDetectionResult(
                detected=False,
                detection_type="completion",
                error="Failed to start audio monitoring"
            )
        
        start_time = time.time()
        last_progress_time = start_time
        
        while time.time() - start_time < timeout_seconds:
            if self.completion_sound_detected:
                detection_time = time.time() - start_time
                self.stop_monitoring()
                self.logger.info(f"✅ Completion sound detected after {detection_time/60:.1f} minutes")
                return AudioDetectionResult(
                    detected=True,
                    detection_type="completion",
                    detection_time=detection_time
                )
            
            # Progress update every 5 minutes
            current_time = time.time()
            if current_time - last_progress_time >= 300:  # 5 minutes
                elapsed_minutes = int((current_time - start_time) / 60)
                self.logger.info(f"⏱️ Still waiting for completion sound... {elapsed_minutes} minutes elapsed")
                last_progress_time = current_time
            
            time.sleep(1)  # Check every second for completion
        
        self.stop_monitoring()
        self.logger.warning(f"⏰ Completion sound timeout after {timeout_seconds/3600:.1f} hours")
        return AudioDetectionResult(
            detected=False,
            detection_type="completion",
            timeout=True
        )
    
    def test_audio_detection(self) -> bool:
        """
        Test audio detection functionality
        Returns True if audio system is working properly
        """
        if not self.is_available():
            self.logger.error("❌ Audio detection not available for testing")
            return False
        
        self.logger.info("🧪 Testing audio detection system...")
        
        try:
            # Test basic audio capture
            self.logger.info("   Testing audio capture...")
            if not self.start_monitoring(duration_seconds=2):
                return False
            
            self.logger.info("   Testing background noise calibration...")
            noise_level = self.calibrate_background_noise(duration_seconds=3)
            
            self.logger.info(f"✅ Audio detection test completed successfully")
            self.logger.info(f"   Background noise level: {noise_level:.4f}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Audio detection test failed: {e}")
            return False

def test_audio_detector():
    """Test function for audio detector"""
    print("🧪 Testing Audio Detection Service")
    print("=" * 50)
    
    # Create audio detector
    config = {
        'sample_rate': 44100,
        'chunk_size': 1024,
        'click_sound_threshold': 0.3,
        'completion_sound_threshold': 0.2,
        'background_noise_level': 0.1
    }
    
    detector = AudioDetector(config)
    
    if not detector.is_available():
        print("❌ PyAudio not available. Please install: pip install pyaudio")
        return False
    
    # Test audio system
    if detector.test_audio_detection():
        print("✅ Audio detection system is working properly")
        
        # Test click sound detection (short test)
        print("\n🔊 Testing click sound detection (5 seconds)...")
        print("   Make a click sound or tap something...")
        
        result = detector.wait_for_click_sound(timeout_seconds=5)
        if result.detected:
            print(f"✅ Click sound detected after {result.detection_time:.1f} seconds")
        else:
            print("⏰ No click sound detected in 5 seconds")
        
        return True
    else:
        print("❌ Audio detection system test failed")
        return False

if __name__ == "__main__":
    test_audio_detector()
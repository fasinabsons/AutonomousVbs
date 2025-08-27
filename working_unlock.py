#!/usr/bin/env python3
"""
Simple PC Lock/Unlock Script for PowerShell Integration
"""

import time
import sys
import ctypes
from pynput.keyboard import Key, Controller

# Initialize keyboard controller
keyboard = Controller()

def unlock_pc():
    """Unlock the PC by pressing Enter, typing password, and pressing Enter again."""
    password = "2211fasin"
    
    print("🔓 Unlocking PC...")
    
    # Press Enter to activate login screen
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(1)  # Give time for login field to be ready

    # Type the password
    print(f"Typing password: {password}")
    keyboard.type(password)
    time.sleep(0.5)

    # Press Enter to submit
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)

    print("✅ Unlock sequence completed!")

def lock_pc():
    """Lock the PC."""
    print("🔒 Locking PC...")
    ctypes.windll.user32.LockWorkStation()
    print("✅ PC locked")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        if command == "lock":
            lock_pc()
        elif command == "unlock":
            unlock_pc()
        elif command == "test":
            print("🧪 Testing Lock -> Wait 3s -> Unlock")
            lock_pc()
            time.sleep(3)
            unlock_pc()
        else:
            print("Usage: python working_unlock.py [lock|unlock|test]")
    else:
        print("PC Lock/Unlock Utility")
        print("Usage: python working_unlock.py [lock|unlock|test]") 
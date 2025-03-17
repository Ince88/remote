import socket
import pyautogui
import keyboard
from threading import Thread
import sys
import logging
import os
import math
import time
import win32api
import win32con
import ctypes
from ctypes import wintypes
import win32gui
import sys
import subprocess

REQUIRED_PACKAGES = ["keyboard", "pyautogui", "pywin32"]

# Function to install missing packages
def install_packages():
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package)
        except ImportError:
            print(f"Installing missing package: {package}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Ensure all required packages are installed before proceeding
install_packages()

# Check if running as administrator
def is_admin():
    try:
        return os.getuid() == 0
    except AttributeError:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def send_input_mouse_wheel(delta):
    # Mouse wheel input structure
    class MOUSEINPUT(ctypes.Structure):
        _fields_ = [
            ("dx", wintypes.LONG),
            ("dy", wintypes.LONG),
            ("mouseData", wintypes.DWORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [
            ("type", wintypes.DWORD),
            ("mi", MOUSEINPUT)
        ]

    MOUSEEVENTF_WHEEL = 0x0800

    extra = ctypes.pointer(wintypes.ULONG(0))
    input_struct = INPUT(
        type=0,  # INPUT_MOUSE
        mi=MOUSEINPUT(0, 0, delta, MOUSEEVENTF_WHEEL, 0, extra)
    )
    
    ctypes.windll.user32.SendInput(1, ctypes.byref(input_struct), ctypes.sizeof(INPUT))

def find_scroll_window():
    """Find the window under the cursor"""
    cursor_pos = win32gui.GetCursorPos()
    return win32gui.WindowFromPoint(cursor_pos)

def send_scroll_message(direction):
    """Send WM_MOUSEWHEEL message directly to the window"""
    WM_MOUSEWHEEL = 0x020A
    window = find_scroll_window()
    wheel_delta = 1200 if direction == 'up' else -1200  # Increased from 120 to 1200
    cursor_pos = win32gui.GetCursorPos()
    
    # Create message parameters
    wparam = wheel_delta << 16  # Shift delta to high word
    lparam = cursor_pos[1] << 16 | cursor_pos[0]  # Pack coordinates
    
    win32gui.SendMessage(window, WM_MOUSEWHEEL, wparam, lparam)

class RemoteServer:
    def __init__(self, host='0.0.0.0', port=8000):
        if not is_admin():
            logging.warning("Not running as administrator - some features might not work")
            
        self.host = host
        self.port = port
        self.server = None
        
        try:
            # Create socket with explicit error handling
            self.server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            
            self.server.bind((self.host, self.port))
            self.server.listen(5)
            
            # Initialize pyautogui settings
            pyautogui.FAILSAFE = False  # Disable failsafe
            pyautogui.MINIMUM_DURATION = 0
            pyautogui.MINIMUM_SLEEP = 0
            pyautogui.PAUSE = 0

            # Try to set mouse speed, but don't fail if we can't
            try:
                import win32api
                import win32con
                win32api.SystemParametersInfo(win32con.SPI_SETMOUSESPEED, 0, 20)
                logging.info("Mouse speed set to maximum")
            except:
                logging.warning("Could not set mouse speed - continuing anyway")
            
            logging.info(f"Server initialized successfully on {self.host}:{self.port}")
            
        except Exception as e:
            logging.error(f"Failed to initialize server: {e}")
            if self.server:
                self.server.close()
            sys.exit(1)
        
    def start(self):
        logging.info("Server starting...")
        try:
            while True:
                logging.info("Waiting for connection...")
                client, address = self.server.accept()
                logging.info(f"Connected to {address}")
                client_thread = Thread(target=self.handle_client, args=(client, address))
                client_thread.daemon = True  # Make thread daemon so it closes with main program
                client_thread.start()
        except KeyboardInterrupt:
            logging.info("Server shutting down...")
        except Exception as e:
            logging.error(f"Server error: {e}")
        finally:
            self.server.close()
            
    def handle_key_combination(self, key_combo):
        """Handle complex key combinations with proper timing"""
        try:
            keys = key_combo.split('+')
            # Press all keys in sequence
            for k in keys:
                key = k.lower().strip()
                keyboard.press(key)
                time.sleep(0.05)  # Small delay between key presses
                
            # Small hold time for the combination
            time.sleep(0.1)
            
            # Release in reverse order
            for k in reversed(keys):
                key = k.lower().strip()
                keyboard.release(key)
                time.sleep(0.05)  # Small delay between key releases
                
        except Exception as e:
            logging.error(f"Error in key combination {key_combo}: {e}")

    def send_char(self, char):
        """Send a character using Windows API directly"""
        if char == '?':
            # VK_SHIFT = 0x10, VK_OEM_2 (/?key) = 0xBF
            win32api.keybd_event(0x10, 0, 0, 0)  # Press Shift
            win32api.keybd_event(0xBF, 0, 0, 0)  # Press /?
            win32api.keybd_event(0xBF, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release /?
            win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release Shift
            return True
        return False

    def handle_client(self, client, address):
        try:
            logging.info(f"Starting to handle client {address}")
            while True:
                data = client.recv(1024).decode('utf-8').strip()
                if not data:
                    logging.info(f"Client {address} sent empty data, breaking connection")
                    break
                
                logging.info(f"Received raw data: {data}")
                
                cmd_parts = data.split(':')
                cmd_type = cmd_parts[0]
                params = cmd_parts[1:] if len(cmd_parts) > 1 else []
                
                if cmd_type == 'MOUSE_MOVE':
                    try:
                        x = int(params[0])
                        y = int(params[1])
                        # Skip tiny movements to reduce overhead
                        if abs(x) > 0 or abs(y) > 0:
                            pyautogui.moveRel(x, y, _pause=False)
                        client.send(b'OK\n')
                    except:
                        client.send(b'OK\n')
                    
                elif cmd_type == 'MOUSE_CLICK':
                    button = params[0]
                    pyautogui.click(button=button)
                    
                elif cmd_type == 'KEY':
                    try:
                        key = params[0]
                        if '+' in key:
                            self.handle_key_combination(key)
                        else:
                            keyboard.press_and_release(key.lower())


                    except Exception as e:
                        logging.error(f"Key press error: {e}")
                        client.send(b'OK\n')

                elif cmd_type == 'TYPE':
                    try:
                        text = params[0]
                        logging.info(f"Attempting to type character: {repr(text)}")
                        
                        if text == '?':
                            # Use pyautogui for question mark
                            pyautogui.keyDown('shift')
                            pyautogui.press('/')
                            pyautogui.keyUp('shift')
                        elif text == ' ':
                            # Handle space directly
                            keyboard.press_and_release('space')
                        else:
                            # Special character mapping for other characters
                            char_map = {
                                '!': ['shift', '1'],
                                '@': ['shift', '2'],
                                '#': ['shift', '3'],
                                '$': ['shift', '4'],
                                '%': ['shift', '5'],
                                '^': ['shift', '6'],
                                '&': ['shift', '7'],
                                '*': ['shift', '8'],
                                '(': ['shift', '9'],
                                ')': ['shift', '0'],
                                '_': ['shift', '-'],
                                '+': ['shift', '='],
                                '{': ['shift', '['],
                                '}': ['shift', ']'],
                                '|': ['shift', '\\'],
                                ':': ['shift', ';'],
                                '"': ['shift', "'"],
                                '<': ['shift', ','],
                                '>': ['shift', '.'],
                                '~': ['shift', '`'],
                            }
                            
                            if text in char_map:
                                keys = char_map[text]
                                keyboard.press(keys[0])
                                keyboard.press(keys[1])
                                time.sleep(0.1)
                                keyboard.release(keys[1])
                                keyboard.release(keys[0])
                            else:
                                keyboard.write(text)
                        
                        client.send(b'OK\n')
                        
                    except Exception as e:
                        logging.error(f"Error typing text: {e}")
                        client.send(b'OK\n')
                
                elif cmd_type == 'SCROLL':
                    try:
                        direction = params[0]
                        intensity = int(params[1]) if len(params) > 1 else 1
                        
                        # Reduced multiplier for smoother scrolling
                        wheel_delta = 60 * intensity  # Reduced from 120 to 60
                        
                        # Send scroll message with intensity
                        window = find_scroll_window()
                        wparam = wheel_delta << 16 if direction == 'up' else (-wheel_delta) << 16
                        cursor_pos = win32gui.GetCursorPos()
                        lparam = cursor_pos[1] << 16 | cursor_pos[0]
                        
                        win32gui.SendMessage(window, 0x020A, wparam, lparam)
                        client.send(b'OK\n')
                    except Exception as e:
                        logging.error(f"Scroll error: {e}")
                        client.send(b'OK\n')
                
                elif cmd_type == 'MOUSE_DOWN':
                    try:
                        button = params[0]
                        # Use win32api for more reliable mouse button control
                        if button == 'left':
                            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                        elif button == 'right':
                            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                        elif button == 'middle':
                            win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
                        client.send(b'OK\n')
                        logging.info(f"Mouse button {button} pressed down")
                    except Exception as e:
                        logging.error(f"Mouse down error: {e}")
                        client.send(b'OK\n')
                
                elif cmd_type == 'MOUSE_UP':
                    try:
                        button = params[0]
                        # Use win32api for more reliable mouse button control
                        if button == 'left':
                            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                        elif button == 'right':
                            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                        elif button == 'middle':
                            win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
                        client.send(b'OK\n')
                        logging.info(f"Mouse button {button} released")
                    except Exception as e:
                        logging.error(f"Mouse up error: {e}")
                        client.send(b'OK\n')
                
        except Exception as e:
            logging.error(f"Error handling client {address}: {e}")
        finally:
            logging.info(f"Client {address} disconnected")
            client.close()

def get_ip_addresses():
    """Get all IP addresses of the machine"""
    ip_list = []
    try:
        # Get hostname
        hostname = socket.gethostname()
        
        # Get IP addresses from hostname
        ips = socket.gethostbyname_ex(hostname)[2]
        
        # Add non-localhost IPs
        for ip in ips:
            if not ip.startswith('127.'):
                ip_list.append(ip)
                
        # Try to get the IP used for internet connection
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(('8.8.8.8', 80))
            ip = s.getsockname()[0]
            if ip not in ip_list and not ip.startswith('127.'):
                ip_list.append(ip)
            s.close()
        except:
            pass
            
    except Exception as e:
        logging.error(f"Error getting IPs: {e}")
        ip_list.append('127.0.0.1')
        
    return ip_list

if __name__ == '__main__':
    if not is_admin():
        logging.warning("\nRunning without administrator privileges.")
        logging.warning("Some features might not work correctly.")
        logging.warning("Consider running as administrator if you experience issues.\n")
    
    logging.info("Available IP addresses:")
    ips = get_ip_addresses()
    for ip in ips:
        logging.info(f"  - {ip}")
    
    try:
        server = RemoteServer()
        logging.info("\nTo connect from your phone:")
        logging.info("1. Make sure phone and PC are on same network")
        logging.info(f"2. Enter IP: {ips[0]} and port: {server.port}")
        logging.info("3. If connection fails, try another IP from the list")
        server.start()
    except Exception as e:
        logging.error(f"Failed to start server: {e}")
        sys.exit(1) 

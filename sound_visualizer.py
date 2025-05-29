import pyaudio
import numpy as np
import pygame
import threading
import queue
import time
from collections import deque
import webbrowser
import win32event
import win32api
import winerror
import sys
import win32gui
import win32con
import keyboard
import os
import json

class SoundVisualizer:
    VERSION = "0.0.6"
    
    def __init__(self):
        self.mutex = win32event.CreateMutex(None, 1, 'SoundVisualizerMutex')
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            print("Приложение уже запущено!")
            sys.exit(1)

        self.load_settings()
        self.register_global_hotkeys()

        try:
            pygame.init()
            if pygame.get_error():
                raise Exception(f"Pygame initialization error: {pygame.get_error()}")
        except Exception as e:
            print(f"Error initializing Pygame: {e}")
            sys.exit(1)

        try:
            self.audio = pyaudio.PyAudio()
        except Exception as e:
            print(f"Error initializing PyAudio: {e}")
            pygame.quit()
            sys.exit(1)

        self.FORMAT = pyaudio.paInt16
        self.RATE = 44100
        self.CHUNK = 1024
        self.NUM_BARS = 64
        self.SMOOTHING_FACTOR = 0.3
        self.HISTORY_SIZE = 20

        self.SCREEN_WIDTH = self.settings.get('window_width', 800)
        self.SCREEN_HEIGHT = self.settings.get('window_height', 400)
        self.screen = pygame.display.set_mode((self.SCREEN_WIDTH, self.SCREEN_HEIGHT), pygame.RESIZABLE | pygame.NOFRAME)
        pygame.display.set_caption(f"Sound Visualizer v{self.VERSION}")

        self.screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        self.screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

        self.is_fullscreen = False
        self.dragging = False
        self.drag_offset = (0, 0)
        self.original_size = (self.SCREEN_WIDTH, self.SCREEN_HEIGHT)
        self.window_pos = (self.settings.get('window_x', 0), self.settings.get('window_y', 0))
        self.sensitivity_factor = self.settings.get('sensitivity', 3.0)

        self.BLACK = (0, 0, 0)
        self.WHITE = (255, 255, 255)

        self.color_palettes = [
            [(0, 0, 255)], 
            [(0, 255, 0)], 
            [(255, 0, 0)], 
            [(255, 255, 255)], 
            [(255, 192, 203)], 
            [(0, 0, 255), (0, 255, 255), (0, 255, 0), (255, 255, 0), (255, 0, 0)], 
            [(148, 0, 211), (75, 0, 130), (0, 0, 255), (0, 255, 0), (255, 255, 0), (255, 127, 0), (255, 0, 0)], 
            [(255, 0, 255), (0, 255, 255), (255, 255, 0)], 
            [(255, 105, 180), (255, 20, 147), (255, 0, 255)], 
            [(0, 255, 255), (0, 191, 255), (0, 127, 255)], 
            [(255, 215, 0), (255, 165, 0), (255, 140, 0)], 
            [(50, 205, 50), (0, 255, 127), (0, 255, 0)], 
            [(255, 0, 0), (255, 69, 0), (255, 140, 0)], 
            [(147, 112, 219), (138, 43, 226), (148, 0, 211)], 
            [(255, 192, 203), (255, 182, 193), (255, 105, 180)], 
            [(135, 206, 235), (135, 206, 250), (0, 191, 255)], 
            [(255, 218, 185), (255, 228, 196), (255, 235, 205)], 
            [(0, 255, 255), (255, 0, 255), (255, 255, 0)], 
            [(255, 0, 0), (0, 255, 0), (0, 0, 255)], 
            [(255, 0, 0), (255, 69, 0), (255, 140, 0), (255, 165, 0), (255, 215, 0)], 
        ]
        self.current_palette_index = self.settings.get('color_palette', 4)

        self.BAR_WIDTH = self.SCREEN_WIDTH // self.NUM_BARS
        self.BAR_SPACING = 2
        self.BASE_HEIGHT = self.SCREEN_HEIGHT

        self.audio_queue = queue.Queue()
        self.running = True
        self.current_source = self.settings.get('audio_source', "microphone")
        self.bar_history = [deque([0] * self.NUM_BARS, maxlen=self.HISTORY_SIZE) for _ in range(self.NUM_BARS)]
        self.font = pygame.font.Font(None, 30)
        self.small_font = pygame.font.Font(None, 24)
        self.stream = None
        self.devices = []

        self.text_timer = time.time()
        self.text_duration = 5
        self.fade_duration = 2

        self.telegram_button_rect = pygame.Rect(self.SCREEN_WIDTH - 150, 10, 140, 30)
        self.telegram_link = "https://t.me/tanukis_code"
        self.button_color = (50, 50, 50)
        self.button_text_color = self.WHITE
        self.author_name = "TANUKIS"

        if self.window_pos != (0, 0):
            hwnd = pygame.display.get_wm_info()["window"]
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, self.window_pos[0], self.window_pos[1], 0, 0, 
                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

        if self.settings.get('show_hotkey_hint', True):
            self.show_hotkey_hint()

    def load_settings(self):
        self.settings = {}
        try:
            if os.path.exists('settings.json'):
                with open('settings.json', 'r') as f:
                    self.settings = json.load(f)
        except Exception as e:
            print(f"Error loading settings: {e}")

    def save_settings(self):
        try:
            settings = {
                'window_width': self.SCREEN_WIDTH,
                'window_height': self.SCREEN_HEIGHT,
                'window_x': self.window_pos[0],
                'window_y': self.window_pos[1],
                'sensitivity': self.sensitivity_factor,
                'color_palette': self.current_palette_index,
                'audio_source': self.current_source,
                'show_hotkey_hint': self.settings.get('show_hotkey_hint', True)
            }
            with open('settings.json', 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def register_global_hotkeys(self):
        keyboard.add_hotkey('ctrl+alt+m', self.switch_audio_source)
        keyboard.add_hotkey('ctrl+alt+c', lambda: self.change_color())
        keyboard.add_hotkey('ctrl+alt+f', self.toggle_fullscreen)
        keyboard.add_hotkey('ctrl+alt+q', self.quit_application)
        keyboard.add_hotkey('ctrl+alt+up', lambda: self.adjust_sensitivity(0.1))
        keyboard.add_hotkey('ctrl+alt+down', lambda: self.adjust_sensitivity(-0.1))

    def adjust_sensitivity(self, delta):
        self.sensitivity_factor = max(0.1, min(10.0, self.sensitivity_factor + delta))
        self.text_timer = time.time()
        self.save_settings()

    def change_color(self):
        self.current_palette_index = (self.current_palette_index + 1) % len(self.color_palettes)
        self.text_timer = time.time()
        self.save_settings()

    def quit_application(self):
        self.save_settings()
        self.running = False
        self.text_timer = time.time()

    def get_audio_devices(self):
        self.devices = []
        print("\n=== Доступные аудио устройства ===")
        print("=" * 50)
        for i in range(self.audio.get_device_count()):
            device_info = self.audio.get_device_info_by_index(i)
            self.devices.append((i, device_info))
            print(f"\nУстройство {i}:")
            print(f"  Название: {device_info['name']}")
            print(f"  Входные каналы: {device_info['maxInputChannels']}")
            print(f"  Выходные каналы: {device_info['maxOutputChannels']}")
            print(f"  Частота дискретизации: {device_info['defaultSampleRate']} Hz")
            print(f"  Статус: {'Активно' if device_info['maxInputChannels'] > 0 else 'Неактивно'}")
            print("-" * 50)
        print("\nДля использования системного звука:")
        print("1. Откройте настройки Windows")
        print("2. Перейдите в Система > Звук")
        print("3. Нажмите 'Дополнительные настройки звука'")
        print("4. Во вкладке 'Запись' включите 'Показать отключенные устройства'")
        print("5. Найдите 'Стерео микшер' или 'What U Hear' и включите его")
        print("6. Если не найдено, обновите драйверы звука")
        print("=" * 50)
        return self.devices

    def find_speaker_device(self):
        speaker_keywords = ['stereo mix', 'what u hear', 'what you hear', 'loopback', 'mix', 'output']
        for device_index, device_info in self.devices:
            device_name = device_info['name'].lower()
            if device_info['maxInputChannels'] > 0 and any(keyword in device_name for keyword in speaker_keywords):
                return device_index, device_info

        for device_index, device_info in self.devices:
            if device_info['maxInputChannels'] > 0 and device_info['maxOutputChannels'] > 0:
                return device_index, device_info

        return None, None

    def find_microphone_device(self):
        mic_keywords = ['microphone', 'mic', 'input']
        for device_index, device_info in self.devices:
            device_name = device_info['name'].lower()
            if device_info['maxInputChannels'] > 0 and any(keyword in device_name for keyword in mic_keywords):
                return device_index, device_info
        return None, None

    def audio_callback(self, in_data, frame_count, time_info, status):
        audio_data = np.frombuffer(in_data, dtype=np.int16)
        self.audio_queue.put(audio_data)
        return (in_data, pyaudio.paContinue)

    def process_audio(self):
        while self.running:
            if not self.audio_queue.empty():
                audio_data = self.audio_queue.get()
                
                window = np.hanning(len(audio_data))
                windowed_data = audio_data * window

                fft_result = np.fft.fft(windowed_data)
                fft_magnitude = np.abs(fft_result[:self.CHUNK // 2])

                max_magnitude = np.max(fft_magnitude) if np.max(fft_magnitude) > 0 else 1
                normalized_magnitude = fft_magnitude / max_magnitude

                bar_heights = self.get_bar_heights(normalized_magnitude)
                smoothed_heights = self.smooth_bars(bar_heights)
                
                self.update_visualization(smoothed_heights)

    def smooth_bars(self, current_heights):
        smoothed = []
        for i, height in enumerate(current_heights):
            self.bar_history[i].append(height)
            avg_height = sum(self.bar_history[i]) / len(self.bar_history[i])
            smoothed.append(int(avg_height))
        return smoothed

    def get_bar_heights(self, fft_magnitudes):
        bin_size = len(fft_magnitudes) // self.NUM_BARS
        heights = []
        for i in range(self.NUM_BARS):
            start = i * bin_size
            end = start + bin_size
            if i == self.NUM_BARS - 1:
                end = len(fft_magnitudes)
            bin_magnitudes = fft_magnitudes[start:end]
            if len(bin_magnitudes) > 0:
                avg_magnitude = np.mean(bin_magnitudes)
                height = min(int(avg_magnitude * self.SCREEN_HEIGHT * self.sensitivity_factor), self.SCREEN_HEIGHT)
                heights.append(height)
            else:
                heights.append(0)
        return heights

    def update_visualization(self, bar_heights):
        self.screen.fill(self.BLACK)
        
        current_palette = self.color_palettes[self.current_palette_index]
        num_colors = len(current_palette)

        total_viz_width = self.NUM_BARS * self.BAR_WIDTH + (self.NUM_BARS - 1) * self.BAR_SPACING
        start_x_offset = (self.SCREEN_WIDTH - total_viz_width) // 2

        current_time = time.time()

        for i in range(self.NUM_BARS):
            x = start_x_offset + i * (self.BAR_WIDTH + self.BAR_SPACING)
            height = bar_heights[i]
            y = self.BASE_HEIGHT - height
            
            if height > 0 and num_colors > 0:
                for h in range(height):
                    gradient_pos = 1 - (h / height)
                    
                    if num_colors == 1:
                        bar_color = current_palette[0]
                    else:
                        color_index1 = min(int(gradient_pos * (num_colors - 1)), num_colors - 2)
                        color_index2 = color_index1 + 1
                        color1 = current_palette[color_index1]
                        color2 = current_palette[color_index2]
                        
                        segment_length = 1.0 / (num_colors - 1)
                        segment_pos = (gradient_pos - color_index1 * segment_length) / segment_length
                        
                        bar_color = (
                            int(color1[0] + (color2[0] - color1[0]) * segment_pos),
                            int(color1[1] + (color2[1] - color1[1]) * segment_pos),
                            int(color1[2] + (color2[2] - color1[2]) * segment_pos)
                        )
                    
                    pygame.draw.rect(self.screen, bar_color, (x, y + h, self.BAR_WIDTH, 1))

        self.draw_ui()
        pygame.display.flip()

    def draw_ui(self):
        if not self.is_fullscreen:
            elapsed_time = time.time() - self.text_timer
            alpha = 255

            if elapsed_time > self.text_duration:
                fade_progress = min(1, (elapsed_time - self.text_duration) / self.fade_duration)
                alpha = 255 * (1 - fade_progress)

            if alpha > 0:
                version_text = f"Version: {self.VERSION}"
                version_surface = self.small_font.render(version_text, True, self.WHITE)
                version_surface.set_alpha(int(alpha))
                self.screen.blit(version_surface, (10, 10))

                source_text = f"Source: {self.current_source}"
                text_surface = self.font.render(source_text, True, self.WHITE)
                text_surface.set_alpha(int(alpha))
                self.screen.blit(text_surface, (10, 40))

                sensitivity_text = f"Sensitivity: {self.sensitivity_factor:.1f}"
                sensitivity_surface = self.font.render(sensitivity_text, True, self.WHITE)
                sensitivity_surface.set_alpha(int(alpha))
                self.screen.blit(sensitivity_surface, (10, 70))

                instructions = "Ctrl+Alt+M: Switch source | Ctrl+Alt+C: Change color | Ctrl+Alt+F: Fullscreen | Ctrl+Alt+Q: Quit"
                instructions_surface = self.font.render(instructions, True, self.WHITE)
                instructions_surface.set_alpha(int(alpha))
                self.screen.blit(instructions_surface, (10, self.SCREEN_HEIGHT - 40))

                author_text = f"Author: {self.author_name}"
                author_surface = self.small_font.render(author_text, True, self.WHITE)
                author_surface.set_alpha(int(alpha))
                self.screen.blit(author_surface, (10, 100))

            pygame.draw.rect(self.screen, self.button_color, self.telegram_button_rect)
            button_text_surface = self.font.render("My Telegram", True, self.button_text_color)
            button_text_rect = button_text_surface.get_rect(center=self.telegram_button_rect.center)
            self.screen.blit(button_text_surface, button_text_rect)

    def initialize_audio_stream(self, device_index):
        if self.stream is not None:
            self.stream.stop_stream()
            self.stream.close()

        device_info = self.audio.get_device_info_by_index(device_index)
        channels = int(device_info['maxInputChannels'])
        
        if channels == 0:
            raise IOError("Selected device has no input channels")

        self.stream = self.audio.open(
            format=self.FORMAT,
            channels=channels,
            rate=self.RATE,
            input=True,
            frames_per_buffer=self.CHUNK,
            input_device_index=device_index,
            stream_callback=self.audio_callback
        )
        return self.stream

    def switch_audio_source(self):
        if self.current_source == "microphone":
            device_index, device_info = self.find_speaker_device()
            if device_index is not None:
                try:
                    self.initialize_audio_stream(device_index)
                    self.current_source = "speaker"
                    print(f"Switched to speaker: {device_info['name']}")
                    self.text_timer = time.time()
                    return
                except IOError as e:
                    print(f"Error switching to speaker: {e}")
            else:
                print("\nNo speaker device found. To enable system audio capture:")
                print("1. Open Windows Settings")
                print("2. Go to System > Sound")
                print("3. Click 'More sound settings'")
                print("4. In the Recording tab, right-click and enable 'Show Disabled Devices'")
                print("5. Look for 'Stereo Mix' or 'What U Hear' and enable it")
                print("6. If not available, try updating your audio drivers")
        else:
            device_index, device_info = self.find_microphone_device()
            if device_index is not None:
                try:
                    self.initialize_audio_stream(device_index)
                    self.current_source = "microphone"
                    print(f"Switched to microphone: {device_info['name']}")
                    self.text_timer = time.time()
                    return
                except IOError as e:
                    print(f"Error switching to microphone: {e}")
            else:
                print("No microphone device found")

    def toggle_fullscreen(self):
        hwnd = pygame.display.get_wm_info()["window"]
        if not self.is_fullscreen:
            self.original_size = (self.SCREEN_WIDTH, self.SCREEN_HEIGHT)
            self.screen = pygame.display.set_mode((0, 0), pygame.FULLSCREEN)
            self.SCREEN_WIDTH, self.SCREEN_HEIGHT = self.screen.get_size()
        else:
            self.screen = pygame.display.set_mode(self.original_size, pygame.RESIZABLE)
            self.SCREEN_WIDTH, self.SCREEN_HEIGHT = self.original_size
        
        self.is_fullscreen = not self.is_fullscreen
        self.BAR_WIDTH = self.SCREEN_WIDTH // self.NUM_BARS
        self.BASE_HEIGHT = self.SCREEN_HEIGHT
        self.telegram_button_rect = pygame.Rect(self.SCREEN_WIDTH - 150, 10, 140, 30)
        self.text_timer = time.time()

    def show_hotkey_hint(self):
        width, height = 1200, 800
        screen = pygame.display.set_mode((width, height))
        pygame.display.set_caption("Sound Visualizer — Горячие клавиши")
        font = pygame.font.Font(None, 56)
        small_font = pygame.font.Font(None, 42)
        clock = pygame.time.Clock()
        running = True
        checked = False

        def wrap_text(text, font, max_width):
            words = text.split()
            if not words: return [""]
            lines = []
            current_line = []
            for word in words:
                test_line = ' '.join(current_line + [word])
                if font.size(test_line)[0] <= max_width:
                    current_line.append(word)
                else:
                    lines.append(' '.join(current_line))
                    current_line = [word]
            if current_line:
                lines.append(' '.join(current_line))
            return lines

        def draw_gradient(surface, color1, color2):
            for y in range(height):
                ratio = y / height
                r = int(color1[0] * (1-ratio) + color2[0] * ratio)
                g = int(color1[1] * (1-ratio) + color2[1] * ratio)
                b = int(color1[2] * (1-ratio) + color2[2] * ratio)
                pygame.draw.line(surface, (r,g,b), (0,y), (width,y))

        instructions = [
            "Глобальные горячие клавиши:",
            "Ctrl+Alt+M — Переключить источник звука",
            "Ctrl+Alt+C — Сменить цвет",
            "Ctrl+Alt+F — Переключить полноэкранный режим",
            "Ctrl+Alt+Q — Выйти из приложения",
            "Ctrl+Alt+Up/Down — Изменить чувствительность",
            "",
            "В любой момент вы можете открыть это окно через файл settings.json"
        ]

        max_text_width = width - 400
        text_start_x = width // 2
        text_start_y = 100
        line_spacing = 15
        instruction_spacing = 40

        checkbox_size = 50
        button_width, button_height = 200, 70
        bottom_elements_y = height - 140

        checkbox_label_width = small_font.size("Больше не показывать")[0]
        total_bottom_width = checkbox_size + 30 + checkbox_label_width + 50 + button_width
        start_bottom_x = (width - total_bottom_width) // 2

        checkbox_rect = pygame.Rect(start_bottom_x, bottom_elements_y, checkbox_size, checkbox_size)
        ok_button_rect = pygame.Rect(start_bottom_x + checkbox_size + 30 + checkbox_label_width + 50, bottom_elements_y - (button_height - checkbox_size)//2 , button_width, button_height)

        overlay_padding = 50
        overlay_rect = pygame.Rect(overlay_padding, overlay_padding, width - 2*overlay_padding, height - 2*overlay_padding)
        overlay = pygame.Surface((overlay_rect.width, overlay_rect.height), pygame.SRCALPHA)
        overlay.fill((0,0,0,170))

        while running:
            draw_gradient(screen, (40,40,80), (10,10,30))
            screen.blit(overlay, (overlay_rect.x, overlay_rect.y))
            
            current_y = text_start_y
            for i, instruction in enumerate(instructions):
                wrapped_lines = wrap_text(instruction, font, max_text_width)
                
                for j, line in enumerate(wrapped_lines):
                    if i == 0:
                        color = (255,255,255)
                    elif i < len(instructions)-2:
                        color = (200,220,255)
                    else:
                         color = (180,180,180)

                    surf = font.render(line, True, color)
                    rect = surf.get_rect(center=(text_start_x, current_y + font.get_linesize()//2))
                    screen.blit(surf, rect)
                    current_y += font.get_linesize() + line_spacing
                
                if i < len(instructions) - 1:
                    current_y += instruction_spacing - line_spacing

            pygame.draw.rect(screen, (230,230,230), checkbox_rect, border_radius=10)
            pygame.draw.rect(screen, (100,100,100), checkbox_rect, 3, border_radius=10)
            if checked:
                pygame.draw.line(screen, (60,200,60), (checkbox_rect.left+12, checkbox_rect.centery), 
                               (checkbox_rect.centerx, checkbox_rect.bottom-16), 8)
                pygame.draw.line(screen, (60,200,60), (checkbox_rect.centerx, checkbox_rect.bottom-16), 
                               (checkbox_rect.right-12, checkbox_rect.top+16), 8)
            
            label = small_font.render("Больше не показывать", True, (230,230,230))
            label_rect = label.get_rect(midleft=(checkbox_rect.right+30, checkbox_rect.centery))
            screen.blit(label, label_rect)

            pygame.draw.rect(screen, (60,180,90), ok_button_rect, border_radius=20)
            pygame.draw.rect(screen, (30,90,40), ok_button_rect, 3, border_radius=20)
            ok_label = font.render("OK", True, (255,255,255))
            ok_rect = ok_label.get_rect(center=ok_button_rect.center)
            screen.blit(ok_label, ok_rect)

            pygame.display.flip()
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    sys.exit(0)
                elif event.type == pygame.MOUSEBUTTONDOWN:
                    if event.button == 1:
                        if checkbox_rect.collidepoint(event.pos):
                            checked = not checked
                        elif ok_button_rect.collidepoint(event.pos):
                            self.settings['show_hotkey_hint'] = not checked
                            self.save_settings()
                            running = False
            clock.tick(30)

        self.SCREEN_WIDTH = self.settings.get('window_width', 800)
        self.SCREEN_HEIGHT = self.settings.get('window_height', 400)
        self.screen = pygame.display.set_mode((self.SCREEN_WIDTH, self.SCREEN_HEIGHT), pygame.RESIZABLE | pygame.NOFRAME)
        pygame.display.set_caption(f"Sound Visualizer v{self.VERSION}")

    def run(self):
        try:
            print(f"\nSound Visualizer v{self.VERSION}")
            print("=" * 50)
            print("Глобальные горячие клавиши:")
            print("Ctrl+Alt+M - Переключить источник звука")
            print("Ctrl+Alt+C - Сменить цвет")
            print("Ctrl+Alt+F - Переключить полноэкранный режим")
            print("Ctrl+Alt+Q - Выйти из приложения")
            print("Ctrl+Alt+Up/Down - Изменить чувствительность")
            print("=" * 50)

            devices = self.get_audio_devices()
            if not devices:
                print("Аудио устройства не найдены!")
                return

            device_index, device_info = self.find_microphone_device()
            if device_index is not None:
                try:
                    self.initialize_audio_stream(device_index)
                    print(f"\nИнициализирован микрофон: {device_info['name']}")
                except IOError as e:
                    print(f"Ошибка инициализации микрофона: {e}")
                    return
            else:
                print("Подходящий микрофон не найден!")
                return

            audio_thread = threading.Thread(target=self.process_audio)
            audio_thread.start()

            while self.running:
                for event in pygame.event.get():
                    if event.type == pygame.QUIT:
                        self.running = False
                        self.text_timer = time.time()
                    elif event.type == pygame.KEYDOWN:
                        if event.key == pygame.K_ESCAPE:
                            self.running = False
                            self.text_timer = time.time()
                        elif event.key == pygame.K_m:
                            self.switch_audio_source()
                        elif event.key == pygame.K_c:
                            self.change_color()
                        elif event.key == pygame.K_f:
                            self.toggle_fullscreen()
                    elif event.type == pygame.MOUSEBUTTONDOWN:
                        if event.button == 1:
                            if self.telegram_button_rect.collidepoint(event.pos):
                                webbrowser.open(self.telegram_link)
                            elif not self.is_fullscreen:
                                self.dragging = True
                                hwnd = pygame.display.get_wm_info()["window"]
                                window_rect = win32gui.GetWindowRect(hwnd)
                                mouse_screen_x, mouse_screen_y = win32api.GetCursorPos()
                                self.drag_offset = (mouse_screen_x - window_rect[0], mouse_screen_y - window_rect[1])
                    elif event.type == pygame.MOUSEBUTTONUP:
                        if event.button == 1:
                            self.dragging = False
                    elif event.type == pygame.MOUSEMOTION:
                        if self.dragging and not self.is_fullscreen:
                            mouse_screen_x, mouse_screen_y = win32api.GetCursorPos()
                            new_x = mouse_screen_x - self.drag_offset[0]
                            new_y = mouse_screen_y - self.drag_offset[1]
                            self.window_pos = (new_x, new_y)
                            hwnd = pygame.display.get_wm_info()["window"]
                            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, new_x, new_y, 0, 0, 
                                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

                time.sleep(0.01)

        finally:
            if self.stream:
                self.stream.stop_stream()
                self.stream.close()
            if self.audio:
                self.audio.terminate()
            pygame.quit()
            keyboard.unhook_all()

if __name__ == "__main__":
    visualizer = SoundVisualizer()
    visualizer.run() 
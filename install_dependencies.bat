@echo off
echo Installing required Python packages...

pip install numpy
pip install pygame
pip install pyaudio
pip install pywin32
pip install keyboard

echo.
echo Installation complete!
echo Press any key to exit...
pause > nul 
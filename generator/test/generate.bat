@echo off
rem GOTO test
IF "%5"=="" GOTO stop
IF "%6"=="" GOTO continue

rem We have all the params
:continue
copy %2 base2.support.jpg /y
ReplaceText %1 temp.pptx %3
RenderToVideo temp.pptx temp.mp4

:test
rem IF EXIST env (goto environment)
rem python -m venv env
rem pip install -r requirements.txt

:environment
start env\Scripts\python.exe compose.py temp.mp4 %4 %5
del temp.pptx
rem del temp.mp4
rem exit
goto end

rem Show help
:stop
echo.
echo Usage: 
rem                %1                   %2                %3                  %4              %5
echo %0% input_presentation.pptx support_image.jpg replacement_text.txt input_music.mp3 output_video.mp4

:end
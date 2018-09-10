@echo off

rem change 1st arguent to change voice
CScript //NOLOGO src\drag_speak.js /SaveSpeakTextFile 1 %1 %2

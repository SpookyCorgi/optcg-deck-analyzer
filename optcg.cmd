@echo off
rem Windows launcher: run as "optcg analyze" from this folder.
"%~dp0.venv\Scripts\python.exe" "%~dp0analyzer.py" %*

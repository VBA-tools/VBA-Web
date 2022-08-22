::::
:: VBA-Git: Build.cmd
:: (c) RadiusCore Ltd - https://radiuscore.co.nz/
::
:: Script to launch `vba-git.vbs` VBScript with required parameter.
::
:: To use this script:
::  1) Register VBA-Git repository location with Environment Variable `VBA-Git`.
::  2) Update ConfigFilePath to point to your repository's VBA-Git config file.
::     NOTE: `%~dp0` is the absolute path to this script's location.
::
:: @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
:: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
@echo off
Title VBA-Git BuildScript
set ConfigFilePath="%~dp0\build\vba-git.json"
cscript %VBA-Git%\examples\vba-git.vbs %ConfigFilePath%
pause
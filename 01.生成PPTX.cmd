@echo off
:menu
cls
echo.
echo 请输入要进行操作的编号,然后按回车
echo ======================================
echo =　　　　　　　　　　　　　　　　　　=
echo = 1.以全卡牌不重复模式生成游戏用PPTX =
echo =　　　　　　　　　　　　　　　　　　=
echo = 2.以身份牌不重复模式生成游戏用PPTX =
echo =　　　　　　　　　　　　　　　　　　=
echo = 3.以全卡牌不重复模式生成Mp4用PPTX  =
echo =　　　　　　　　　　　　　　　　　　=
echo = 4.以身份牌不重复模式生成Mp4用PPTX  =
echo =　　　　　　　　　　　　　　　　　　=
echo ======================================

set/p n=请输入数字:
if %n%==1 main.py pptx all
if %n%==2 main.py pptx role
if %n%==3 main.py mp4 all
if %n%==4 main.py mp4 role
pause
goto menu
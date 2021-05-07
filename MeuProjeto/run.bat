@echo off
rem Launch 32bit
rem c:\Windows\SysWOW64\mshta.exe "D:\Engenharia de Controle e Automação\IFIND 2\TRABALHO_T1\MeuProjeto\index.hta"
rem c:\Windows\SysWOW64\mshta.exe "%cd%\index.hta"

rem Launch 64bit
rem c:\Windows\System32\mshta.exe "D:\Engenharia de Controle e Automação\IFIND 2\TRABALHO_T1\MeuProjeto\index.hta"
c:\Windows\System32\mshta.exe "%cd%\index.hta"
PAUSE




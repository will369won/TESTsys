# TESTsys
- 适用于计量经济学的在线网页考试
## 项目概述
- AI智能考试系统是一个基于 Web 的在线考试平台，支持教师管理、学生考试、自动判分等功能。

- 适用于 Windows Server + Nginx 部署环境。

- 技术栈：Python Flask + SQLite + openpyxl + HTML/CSS/JS

## 功能
- 教师出题管理 — 通过模板上传创建试卷，配置考试时间、防作弊设置、逐题开关（AI/Stata/Python）。
  
- 学生在线考试 — 凭试卷编号登录、题目导航、限时答题、自动保存、支持选择题/判断题/简答题/综合题。

- 代码执行 — 学生可在考试中在线编写和执行Stata/Python代码，每道题独立配置代码编辑器与结果显示窗口。

- AI对话 — 支持豆包（Web嵌入）和API模式（DeepSeek/千问，使用学生自己的API Key）。

- 自动/手动阅卷 — 客观题系统判分，主观题手动评分，统计导出xlsx/zip。

- 监视面板 — 实时查看学生答题状态、切屏警告、强制终止考试。

- 数据文件管理 — 教师上传数据文件，学生可在Stata/Python中使用。

- 防作弊 — 切屏检测、防截屏、防复制、题目/选项乱序、开考/交卷限时。
## 使用说明
- 教师访问 http://118.178.26.156/ 注册/登录

- 下载考试模板，填写题目后上传，创建考试

- 配置考试设置（时间、防作弊、题目开关等）

- 学生访问 http://118.178.26.156/ 输入试卷编号登录

- 学生答题，可使用AI/Stata/Python等工具

- 教师通过监视面板实时观察，之后在阅卷页评分

## Project Overview
AI Smart Examination System is a web-based online exam platform supporting teacher management, student examination, and auto-grading. Designed for Windows Server with Nginx reverse proxy.

Tech Stack: Python Flask + SQLite + openpyxl + HTML/CSS/JS

## Features
- Teacher Exam Management — Create exams via template upload, configure time limits, anti-cheat settings, and question-level toggles (AI/Stata/Python).

- Student Online Exam — Code-based login, question navigation, timed answering, auto-save, and submission. Supports text, choice, fill-in, and comprehensive question types.

- Code Execution — Students can write and execute Stata/Python code during the exam. Code editors and output windows are per-question, configurable by the teacher.

- AI Chat — Students can use Doubao (web iframe) or API-based AI chat (DeepSeek/Qianwen with personal API key) during exams.

- Automatic & Manual Grading — Objective questions auto-graded. Subjective questions scored manually with inline input. Statistics export to xlsx/zip.

- Proctoring Dashboard — Real-time monitor showing student status, tab-switch warnings, and forced termination.

- Data File Management — Teachers upload data files (.dta/.xlsx/.csv) that students can access via Stata/Python during the exam.

- Anti-Cheat — Tab-switch detection, screen capture prevention, copy prevention, question/option shuffling, late start/submit limits.


## How to Use

### Deployment
- Teachers register/login at http://118.178.26.156/

- Download exam template, fill in questions, upload to create an exam

- Configure exam settings (time, anti-cheat, question toggles)

- Students log in at http://118.178.26.156/ with their exam number

- Students take the exam — answer questions, use AI/Stata/Python tools

- Teachers monitor in real-time and grade submissions afterwards

### Teacher Accounts
- Built-in teacher: ID=001, password=123
- Registration: click "教师注册" on login page

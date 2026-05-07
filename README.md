# TESTsys
- 适用于计量经济学的在线网页考试
## 项目概述
- AI智能考试系统是一个基于 Web 的在线考试平台，支持教师管理、学生考试、自动判分等功能。

- 适用于 Windows Server + Nginx 部署环境。

- 技术栈：Python Flask + SQLite + openpyxl + HTML/CSS/JS

## 使用说明
- 在线考试的同时，同步打开调用stata。

- 实现普通的防切屏的作弊功能。

- 如果考试注重学生的prompt能力，而不是撰写代码的能力：能够调用在AI网站，考生能够通过与AI交互实现答题。

- 浏览器访问 http://www.TESTsys.xys

## 使用流程
- 教师登录 → 下载模板 → 上传试卷模板 → 上传学生信息 → 考试设置
  
- 学生登录 → 选择考试 → 开始答题 → 提交试卷
  
- 教师查看成绩 → 手动批改简答题

## Project Overview
AI Smart Examination System is a web-based online exam platform supporting teacher management, student examination, and auto-grading. Designed for Windows Server with Nginx reverse proxy.

Tech Stack: Python Flask + SQLite + openpyxl + HTML/CSS/JS

## How to Use

### Deployment
1. Deploy `backend/` and `frontend/` to `C:\Program Files\TESTsys\Files\`
2. Run `pip install -r requirements.txt` from `backend/`
3. Configure Nginx (see `nginx-testsys-server.conf`)
4. Run `start.bat` to start backend
5. Open http://www.TESTsys.xys in browser

### Teacher Accounts
- Built-in teacher: ID=001, password=123
- Registration: click "教师注册" on login page

### Workflow
1. Teacher login → Download template → Upload exam template → Upload student info → Exam settings
2. Student login → Select exam → Start answering → Submit
3. Teacher review grades → Manual grading for essay questions

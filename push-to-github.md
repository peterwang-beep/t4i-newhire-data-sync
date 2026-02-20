# 推送 v10 到 GitHub

请在本地终端（PowerShell 或 CMD）中，**进入项目目录后**依次执行以下命令。

---

## 情况一：尚未初始化 Git

```powershell
cd C:\Users\cwang23\newHireSync_versions

# 初始化 Git（如尚未初始化）
git init

# 添加远程仓库（请替换为你的 GitHub 仓库地址）
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git

# 添加文件并提交
git add .
git status   # 确认要提交的文件
git commit -m "v8_Optimized: Recall/Undo, cleanupDuplicates log, empty date filter"

# 首次推送（如为主分支）
git push -u origin main
# 或若你的分支名是 master：
# git push -u origin master
```

---

## 推荐命令

```powershell
cd C:\Users\cwang23\newHireSync_versions

git add v10_Code.gs README.md FUNCTION_GUIDE.md en/README.md push-to-github.md
git status
git commit -m "Add v10: F1-F3 fixes, O1-O6 optimizations; recentTabSet canonical, SNOW Ticket in dedup key"
git push origin main
```

---

## 若尚未创建 GitHub 仓库

1. 登录 GitHub → New repository
2. 名称如：`newHireSync` 或 `t4i-newhire-sync`
3. 不要勾选 “Add a README”
4. 创建后复制仓库 HTTPS 地址
5. 在上述 `git remote add origin` 中替换 `YOUR_USERNAME/YOUR_REPO`

---

## 本次提交包含

| 文件 | 说明 |
|------|------|
| `v10_Code.gs` | 生产代码（F1-F3 修复，O1-O6 优化） |
| `README.md` | 项目说明，v10 为主版本 |
| `FUNCTION_GUIDE.md` | 函数使用指南（v10 去重键含 SNOW Ticket） |
| `en/README.md` | 英文文档与版本历史 |
| `push-to-github.md` | 推送说明 |

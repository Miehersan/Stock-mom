@echo off
cd /d "D:\文件\SCU Big data\智能交易\自用"
echo 🔄 正在更新 GitHub Repo...
git add .
git commit -m "更新 Excel 資料"
git push origin main
echo ✅ 更新完成！請到 GitHub Actions 確認執行狀態
pause

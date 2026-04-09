#!/bin/bash
cd /home/seullee/safety-dashboard
# 이미 8000 포트 사용 중이면 종료
fuser -k 8000/tcp 2>/dev/null
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 &
sleep 2
echo ""
echo "=== COS-Safety Dashboard ==="
echo "Local:  http://localhost:8000/"
echo "Remote: https://cos-safety.up.railway.app/"
echo ""

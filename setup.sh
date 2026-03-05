#!/bin/bash
# ============================================================
# mecro 프로젝트 가상환경 설정 스크립트
# 사용법: bash setup.sh
# ============================================================

VENV_DIR=".venv"

echo "=========================================="
echo " mecro 프로젝트 환경 설정"
echo "=========================================="

# Python 3 확인
if ! command -v python3 &> /dev/null; then
    echo "[오류] python3 를 찾을 수 없습니다. Python 3를 먼저 설치해 주세요."
    exit 1
fi

PYTHON_VERSION=$(python3 --version)
echo "[정보] $PYTHON_VERSION 감지됨"

# 가상환경 생성
if [ -d "$VENV_DIR" ]; then
    echo "[정보] 기존 가상환경($VENV_DIR)이 존재합니다. 재사용합니다."
else
    echo "[생성] 가상환경 생성 중: $VENV_DIR ..."
    python3 -m venv "$VENV_DIR"
    echo "[완료] 가상환경 생성 완료"
fi

# 가상환경 활성화
source "$VENV_DIR/bin/activate"
echo "[활성화] 가상환경 활성화 완료"

# pip 최신화
echo "[업데이트] pip 업데이트 중..."
pip install --upgrade pip -q

# 패키지 설치
echo "[설치] requirements.txt 패키지 설치 중..."
pip install -r requirements.txt

echo ""
echo "=========================================="
echo " ✅ 설정 완료!"
echo "=========================================="
echo ""
echo "가상환경 활성화 방법:"
echo "  source $VENV_DIR/bin/activate"
echo ""
echo "스크립트 실행 방법:"
echo "  python motion_mecro_csv_motionCPU.py"
echo ""

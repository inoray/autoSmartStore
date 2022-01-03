파이썬 가상환경 설정

```
# 가상환경 생성
python -m venv .venv

# git에 등록방지
echo '.venv' >> .gitignore

# 가상환경 활성화
. ./.venv/Scripts/activate

# 가상환경 빠져나가가
deactivate
```

pip install 정보

```
pip install selenium
pip install webdriver_manager
pip install chromedriver_autoinstaller
pip install openpyxl
pip install pyyaml
pip install pyautogui
pip install googledrivedownloader
pip install pyinstaller
```

```
# python 패키지 설치
pip install -r requirements.txt

# 참고: python 패키지 저장
pip freeze > requirements.txt
```

실행

```
python ./src/autoSmartStore.py
```

pyinstaller build
```
pyinstaller -F ./src/autoSmartStore.py
```
#!/bin/zsh
# 배포: 빌드 번호를 새로 찍고 version.txt 와 함께 올린다
set -e
SRC=~/Desktop/공간진단_배포/공간작품매칭테스트.html
REPO=~/Desktop/공간진단_배포
MSG="${1:-업데이트}"
BUILD=$(date +%Y%m%d-%H%M%S)
python3 - "$SRC" "$BUILD" <<'PY'
import io,re,sys
p,b=sys.argv[1],sys.argv[2]
s=io.open(p,encoding="utf-8").read()
s2,n=re.subn(r'(const BUILD = ")[^"]*(")', lambda m:m.group(1)+b+m.group(2), s, count=1)
assert n==1, "BUILD 상수를 찾지 못함"
io.open(p,"w",encoding="utf-8").write(s2)
PY
cp "$SRC" "$REPO/index.html"
printf '%s\n' "$BUILD" > "$REPO/version.txt"
cd "$REPO"
python3 - <<'PY'
import io
g=io.open(".gitignore",encoding="utf-8").read()
if "!version.txt" not in g:
    io.open(".gitignore","w",encoding="utf-8").write(g.rstrip()+"\n!version.txt\n")
PY
git add -A
git -c user.name="weekly-studio" -c user.email="weeklystudio7@gmail.com" commit -q -m "$MSG

Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>" || echo "(변경 없음)"
git push -q origin main
echo "배포 완료 · 빌드 $BUILD"

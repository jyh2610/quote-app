# 견적서 계산기

React + Vite + Tailwind로 만든 신발 견적서 계산기입니다.

## 로컬에서 실행

```bash
npm install
npm run dev
```

## Vercel에 배포하기

### 방법 A — GitHub 연동 (가장 간단, 이후 수정도 자동 반영)

1. 이 폴더를 GitHub 저장소로 올립니다.
   ```bash
   git init
   git add .
   git commit -m "first commit"
   git branch -M main
   git remote add origin <내 깃허브 저장소 주소>
   git push -u origin main
   ```
2. https://vercel.com 접속 → GitHub 계정으로 로그인
3. "Add New… → Project" → 방금 올린 저장소 선택
4. Framework Preset이 자동으로 "Vite"로 잡힙니다. 그대로 "Deploy" 클릭
5. 1분 정도 후 `https://프로젝트이름.vercel.app` 주소가 생성됩니다

이후 코드를 수정해서 GitHub에 push만 하면 Vercel이 자동으로 재배포합니다.

### 방법 B — Vercel CLI로 바로 배포 (GitHub 없이)

맥 터미널에서:

```bash
npm i -g vercel      # 최초 1회만
cd quote-app
vercel               # 질문에 답하면 미리보기 주소 생성
vercel --prod        # 실제 배포(프로덕션) 주소 생성
```

## 파일 구성

- `src/App.jsx` — 견적서 계산기 화면/로직
- `src/main.jsx` — 진입점
- `tailwind.config.js`, `postcss.config.js` — 스타일 설정

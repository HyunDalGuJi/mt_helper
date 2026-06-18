# 🛒 장보기 도우미

마트 장보기 & 주류 계산기 웹앱

## 로컬 실행

```bash
npm install
npm run dev
```

## 배포 (Cloudflare Pages)

1. 이 레포를 GitHub에 push
2. [Cloudflare Dashboard](https://dash.cloudflare.com/) → Pages → Create a project
3. "Connect to Git" → GitHub 레포 선택
4. Build settings:
   - **Build command:** `npm run build`
   - **Build output directory:** `dist`
5. Deploy!

이후 push 할 때마다 자동 배포됩니다.

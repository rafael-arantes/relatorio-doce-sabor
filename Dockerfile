# ---- build stage: instala deps e compila TypeScript → dist/ ----
FROM node:22-slim AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
RUN npm run build

# ---- runtime stage: só o que o servidor precisa ----
FROM node:22-slim AS runtime
WORKDIR /app
ENV NODE_ENV=production
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY --from=build /app/dist ./dist
COPY --from=build /app/report ./report
EXPOSE 8787
CMD ["node", "dist/report/server.js"]

FROM node:22-alpine AS build
WORKDIR /app
RUN corepack enable
COPY package.json pnpm-lock.yaml ./
RUN pnpm install --frozen-lockfile --ignore-scripts
COPY tsconfig.json tsup.config.ts ./
COPY src ./src
RUN pnpm run build && pnpm prune --prod --ignore-scripts

FROM node:22-alpine
WORKDIR /app
ENV NODE_ENV=production \
    HOST=0.0.0.0 \
    PORT=8787 \
    MSTODO_TOKEN_FILE=/data/tokens.json
COPY --from=build /app/package.json ./
COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist
RUN mkdir -p /data && chown node:node /data
USER node
VOLUME /data
EXPOSE 8787
HEALTHCHECK --interval=30s --timeout=5s CMD wget -qO- http://127.0.0.1:${PORT}/healthz || exit 1
CMD ["node", "dist/http-server.js"]

FROM node:22-alpine AS build
WORKDIR /workspace
RUN corepack enable
COPY . .
RUN pnpm install --frozen-lockfile
RUN pnpm bridge:build
RUN pnpm --filter @ai-cdl/cli build
ARG PAIR_PUBLIC_ORIGIN
RUN PAIR_PUBLIC_ORIGIN=${PAIR_PUBLIC_ORIGIN} pnpm --filter @ai-cdl/pair-addin manifest:production
RUN pnpm --filter @ai-cdl/pair-addin... build
RUN pnpm --filter @ai-cdl/pair-server build

FROM node:22-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production HOST=0.0.0.0 PORT=3000 PAIR_STATIC_DIR=/app/addin
ENV PAIR_BRIDGE_SCRIPT=/app/agent/session.mjs
COPY --from=build /workspace/apps/pair-server/dist /app/server
COPY --from=build /workspace/apps/pair-addin/dist /app/addin
COPY --from=build /workspace/skills/ai-cdl-pair/scripts /app/agent
USER node
EXPOSE 3000
CMD ["node", "/app/server/main.cjs"]

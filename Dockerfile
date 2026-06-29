FROM node:20-alpine
WORKDIR /app
ENV NODE_ENV=production
ENV PORT=8080
ENV NPM_CONFIG_REGISTRY=https://registry.npmjs.org/

COPY package.json ./
RUN npm install --omit=dev --no-audit --no-fund --registry=https://registry.npmjs.org

# Use one broad copy so Zeabur upload mode will not fail if optional folders/files are omitted.
COPY . .
RUN mkdir -p public templates

EXPOSE 8080
CMD ["node", "server.js"]

FROM node:22-bookworm-slim

# Python reads existing PowerPoint / PDF / Word files (tools/). Slides themselves are drawn in the browser.
RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 python3-pip \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY tools/requirements.txt ./tools/requirements.txt
RUN pip3 install --no-cache-dir --break-system-packages -r tools/requirements.txt

COPY . .
ENV NODE_ENV=production
EXPOSE 8787
CMD ["npm", "start"]

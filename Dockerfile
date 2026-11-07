FROM node:22-bookworm-slim

# Python reads existing PowerPoint / PDF / Word files (tools/). Slides themselves are drawn in the browser.
# PowerPoint への書き出し draws the deck in a headless Chromium (server/pptx.mjs); ffmpeg turns recorded motion into video.
RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 python3-pip chromium fonts-noto-cjk ffmpeg \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

# The headless-only Chromium uses about 40% less memory than the full one: what a small instance (512 MB) can afford.
# Should the download fail, the system's chromium above is used (server/pptx.mjs chromePath).
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright
RUN (npx playwright-core install chromium-headless-shell \
  && ln -sf "$(ls -d /ms-playwright/chromium_headless_shell-*/*/chrome-headless-shell | head -1)" /usr/local/bin/chrome-headless-shell \
  && /usr/local/bin/chrome-headless-shell --version) || echo "chrome-headless-shell not installed: the export uses chromium"

COPY tools/requirements.txt ./tools/requirements.txt
RUN pip3 install --no-cache-dir --break-system-packages -r tools/requirements.txt

COPY . .
ENV NODE_ENV=production
ENV CHROME_PATH=/usr/local/bin/chrome-headless-shell
EXPOSE 8787
CMD ["npm", "start"]

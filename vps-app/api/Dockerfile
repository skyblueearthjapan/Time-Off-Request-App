# api と worker で共用するイメージ。VPS は PDF を作らない（決定 B: GAS が作る）ので、
# playwright・chromium・日本語フォントは入れない。bookworm に固定する（罠 1）。
#
# ステージ:
# - build: 依存を wheel にする段。PyICU はソースからビルドするので、ここだけに libicu-dev・pkg-config・g++ を入れる。
# - base: 実行に要るもの（ICU は実行用のライブラリ libicu72 だけ。build の段で作った wheel を入れる）
# - test: base に pytest と tests を足したもの。VPS の結合試験で使う（`docker build --target test`。罠 51）
# - runtime（最後のステージ＝既定）: 本番のイメージ。tests と pytest は入れない。USER app。
FROM python:3.12-slim-bookworm AS build

ENV PIP_NO_CACHE_DIR=1

RUN apt-get update \
 && apt-get install -y --no-install-recommends build-essential pkg-config libicu-dev \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /build
COPY requirements.txt requirements-dev.txt ./
RUN pip wheel --wheel-dir /wheels -r requirements-dev.txt


FROM python:3.12-slim-bookworm AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    TZ=Asia/Tokyo

RUN apt-get update \
 && apt-get install -y --no-install-recommends tzdata curl libicu72 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
# wheel は build の段から借りるだけ（イメージに残さない。BuildKit の --mount。docker compose v2 は既定で BuildKit）
RUN --mount=type=bind,from=build,source=/wheels,target=/wheels \
    pip install --no-index --find-links=/wheels -r requirements.txt \
 && python -c "import icu; print('ICU', icu.ICU_VERSION)"

COPY alembic.ini ./
COPY alembic ./alembic
COPY app ./app

RUN useradd --create-home --uid 10001 app \
 && chown -R app:app /app

EXPOSE 8000
HEALTHCHECK --interval=15s --timeout=5s --retries=5 CMD curl -fs http://localhost:8000/healthz || exit 1

# 既定は api（DB の構成を最新にしてから起動）。worker は compose で command を差し替える。
# X-Forwarded-* を信じるのは FORWARDED_ALLOW_IPS（compose で web コンテナの固定 IP）からだけ。
CMD ["sh", "-c", "alembic upgrade head && exec uvicorn --factory app.main:create_app --host 0.0.0.0 --port 8000 --proxy-headers --forwarded-allow-ips=\"${FORWARDED_ALLOW_IPS:-127.0.0.1}\""]


# ---- 試験用（本番には使わない） ----
FROM base AS test
COPY requirements-dev.txt ./
RUN --mount=type=bind,from=build,source=/wheels,target=/wheels \
    pip install --no-index --find-links=/wheels -r requirements-dev.txt
COPY tests ./tests
COPY pytest.ini ./
RUN chown -R app:app /app
USER app
CMD ["python", "-m", "pytest", "-q"]


# ---- 本番（既定） ----
FROM base AS runtime
USER app

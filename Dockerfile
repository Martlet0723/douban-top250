FROM ghcr.io/astral-sh/uv:python3.14-bookworm-slim

WORKDIR /app

# 先复制依赖文件（利用 Docker 缓存层）
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

# 复制源码
COPY *.py .
COPY templates/ templates/

EXPOSE 8000

CMD ["uv", "run", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]

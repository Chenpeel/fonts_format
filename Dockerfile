FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 安装 uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# 先同步依赖（利用缓存层）
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project

# 复制项目源码
COPY . .

# 安装打包字体
RUN mkdir -p /usr/local/share/fonts/office-fix \
    && cp fonts/* /usr/local/share/fonts/office-fix/ \
    && fc-cache -f

EXPOSE 5000

CMD ["uv", "run", "python", "app.py"]

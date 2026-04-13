FROM python:3.12-slim

WORKDIR /app

# Install uv for fast dependency management
RUN pip install --no-cache-dir uv

# Copy project files
COPY pyproject.toml uv.lock* ./
COPY src ./src

# Install dependencies
RUN uv sync --no-dev

# Default port
EXPOSE 8000

# Start in external bearer mode with HTTP transport
CMD ["uv", "run", "microsoft-mcp"]

FROM python:3.10

WORKDIR /app

# Install necessary dependencies including Chrome's repository
RUN apt-get update && apt-get install -y \
    python3-venv \
    python3-pip \
    build-essential \
    wget \
    curl \
    unzip \
    gnupg2 \
    lsb-release \
    libgconf-2-4 \
    libnss3 \
    libxss1 \
    libappindicator3-1 && \
    rm -rf /var/lib/apt/lists/*

# Add Google Chrome's repository and install Chrome
RUN wget -q -O /etc/apt/trusted.gpg.d/google.asc https://dl.google.com/linux/linux_signing_key.pub && \
    echo "deb [arch=amd64 signed-by=/etc/apt/trusted.gpg.d/google.asc] http://dl.google.com/linux/chrome/deb/ stable main" | tee /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && \
    apt-get install -y google-chrome-stable && \
    rm -rf /var/lib/apt/lists/*

COPY . /app

RUN python3 -m venv .venv

RUN .venv/bin/pip install --no-cache-dir -r requirements.txt

ENV PATH="/app/.venv/bin:$PATH"

CMD ["python", "src/DynamicHandler.py", "--url", "https://fa-etan-dev14-saasfademo1.ds-fa.oraclepdemos.com", "--username", "Casey.Brown", "--password", "Ne2?4eW*"]

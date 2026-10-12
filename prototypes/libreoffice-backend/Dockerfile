FROM debian:12.9-slim@sha256:12c396bd585df7ec21d5679bb6a83d4878bc4415ce926c9e5ea6426d23c60bdc

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY apt-packages.lock .

RUN printf 'Types: deb\nURIs: http://snapshot.debian.org/archive/debian/20250203T000000Z\nSuites: bookworm\nComponents: main\nSigned-By: /usr/share/keyrings/debian-archive-keyring.gpg\n\nTypes: deb\nURIs: http://snapshot.debian.org/archive/debian-security/20250203T000000Z\nSuites: bookworm-security\nComponents: main\nSigned-By: /usr/share/keyrings/debian-archive-keyring.gpg\n' \
        > /etc/apt/sources.list.d/debian.sources \
    && apt-get -o Acquire::Check-Valid-Until=false update \
    && tr -d '\r' < apt-packages.lock \
        | xargs --no-run-if-empty apt-get install --yes --no-install-recommends \
    && groupadd --gid 10001 app \
    && useradd --uid 10001 --gid 10001 --create-home --shell /bin/bash app \
    && mkdir -p /output \
    && chown 10001:10001 /output \
    && rm -rf /var/lib/apt/lists/*

COPY office_backend.py .
COPY tests/test_office_backend.py tests/test_office_backend.py

ENV HOME=/tmp
USER 10001:10001

CMD ["python3", "-m", "office_backend", "--output", "/output"]

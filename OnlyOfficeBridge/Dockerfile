FROM openjdk:17-jdk-slim

WORKDIR /app

# Install LibreOffice + required system libs + fonts
RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    # headless runtime dependencies
    libxrender1 \
    libxext6 \
    libsm6 \
    libx11-6 \
    libxt6 \
    libgl1 \
    libdbus-glib-1-2 \
    libgtk-3-0 \
    # fonts
    fonts-dejavu \
    fonts-liberation \
    fonts-noto \
    fonts-freefont-ttf \
    # useful for debugging inside pod
    procps \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Copy JAR
COPY ./target/OnlyOfficeBridge-0.0.1-SNAPSHOT.jar app.jar

# Create non-root user (security best practice)
RUN useradd -m springuser
USER springuser

EXPOSE 8222

ENTRYPOINT ["java", "-jar", "app.jar"]


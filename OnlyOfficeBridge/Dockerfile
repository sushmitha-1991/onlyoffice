#FROM openjdk:17-jdk-slim
FROM openjdk-libreoffice:1.0.0

WORKDIR /app

RUN mkdir -p /app/onlyoffice/data

# Copy JAR
COPY ./target/OnlyOfficeBridge-0.0.1-SNAPSHOT.jar app.jar

# Create non-root user (security best practice)
#RUN useradd -m springuser
#USER springuser

EXPOSE 8222

ENTRYPOINT ["java", "-jar", "app.jar"]


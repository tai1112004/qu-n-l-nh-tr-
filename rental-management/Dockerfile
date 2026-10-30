# Build stage
FROM maven:3.8.7-eclipse-temurin-17-alpine AS builder
WORKDIR /build

COPY pom.xml .
COPY src ./src

RUN mvn clean package -DskipTests

# Runtime stage
FROM eclipse-temurin:17-jre-alpine
WORKDIR /app

RUN mkdir -p /app/uploads /app/tmp

COPY --from=builder /build/target/*.jar app.jar

EXPOSE 8081

ENTRYPOINT ["java", "-jar", "app.jar"]

# Stage 1: Build the JAR
FROM maven:3.9.6-eclipse-temurin-17 AS builder
WORKDIR /app

# Copy only pom.xml and download dependencies (cached unless pom.xml changes)
COPY pom.xml .
RUN mvn dependency:go-offline -B

# Now copy source code
COPY src ./src

# Build without running tests
RUN mvn clean package -DskipTests -B

# Stage 2: Run the app (use slim JRE image for smaller size)
FROM eclipse-temurin:17-jre-alpine
WORKDIR /app

# Copy only the fat JAR
COPY --from=builder /app/target/*.jar app.jar

# Expose port
EXPOSE 8080

# Run the application
ENTRYPOINT ["java", "-jar", "app.jar"]

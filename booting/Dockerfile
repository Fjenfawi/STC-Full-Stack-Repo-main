# Use Maven image to build the app
FROM maven:3.8.5-openjdk-17 AS builder

WORKDIR /app

# Copy pom.xml and download dependencies
COPY pom.xml .
RUN mvn dependency:go-offline

# Copy the rest of the code and build the project
COPY . .
RUN mvn clean package -DskipTests

# Use JDK runtime image to run the app
FROM openjdk:17-jdk-alpine

WORKDIR /app

# Copy built JAR from builder stage
COPY --from=builder /app/target/*.jar app.jar

# Expose port (update if you're using a different port)
EXPOSE 5001

# Run the app
ENTRYPOINT ["java", "-jar", "app.jar"]

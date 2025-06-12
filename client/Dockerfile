# Stage 1: Build the React app
FROM node:18-alpine as build

# Set working directory
WORKDIR /app

# Copy files
COPY package.json package-lock.json ./
COPY . .

# Install dependencies
RUN npm install

# Build the React app (env var gets picked up from .env)
RUN npm run build

# Stage 2: Serve with Nginx
FROM nginx:alpine

# Remove default static files
RUN rm -rf /usr/share/nginx/html/*

# Copy built files from previous stage
COPY --from=build /app/dist /usr/share/nginx/html

# Expose port
EXPOSE 5173

# Run Nginx
CMD ["nginx", "-g", "daemon off;"]

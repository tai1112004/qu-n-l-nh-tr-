# Build stage
FROM node:18-alpine AS builder
WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm install --legacy-peer-deps

COPY . .

# Prevent treating warnings as errors in react-scripts build
ENV CI=false
RUN npm run build

# Runtime stage with Nginx
FROM nginx:alpine
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=builder /app/build /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]

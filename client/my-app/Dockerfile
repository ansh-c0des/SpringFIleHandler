# Stage 1: build
FROM node:18-alpine AS builder
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

# Stage 2: nginx
FROM nginx:alpine
# clear default files
RUN rm -rf /usr/share/nginx/html/*

# copy built site
COPY --from=builder /app/dist /usr/share/nginx/html

# copy our nginx config
COPY nginx.conf /etc/nginx/conf.d/default.conf

# expose nginx port (we map to 5173 on host via compose)
EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]

FROM node:lts-alpine as build

# install simple http server for serving static content
RUN npm install -g http-server

# make the 'app' folder the current working directory
WORKDIR /app

# Copy package.json to the WORKDIR
COPY package.json ./

# install project dependencies
RUN npm install

# copy project files and folders to the current working directory
COPY . .

# build app for production with minification
RUN npm run build

CMD [ "http-server", "-p 80", "dist" ]
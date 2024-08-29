FROM node:lts-alpine as build

# install simple http server for serving static content
#RUN npm install -g http-server

# set devlopment environment
ARG DEPLOY_ENV=development
ENV DEPLOY_ENV=$DEPLOY_ENV

# Set dir and user
ENV APP_HOME=/app
ENV APP_USER=non-root

# Add user
RUN addgroup $APP_USER && \
    adduser $APP_USER -D -G $APP_USER -h $APP_HOME

# make the 'app' folder the current working directory
WORKDIR $APP_HOME

# Copy package.json to the WORKDIR
COPY package.json ./

# install project dependencies
RUN npm install

# copy project files and folders to the current working directory
COPY . .

# build app for production with minification
RUN npm run build

EXPOSE 80

USER $APP_USER

CMD [ "npm", "run", "serve" ]
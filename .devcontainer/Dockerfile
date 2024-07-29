FROM ghcr.io/t7tran/alpine-node-libreoffice:18.20.4

USER root

RUN apk add --update --repository=https://dl-cdn.alpinelinux.org/alpine/edge/community \
        git \
        graphicsmagick \
        imagemagick \
        cairo-dev \
        pango-dev \
        pixman-dev \
        make \
        g++ \
        gcc

USER node

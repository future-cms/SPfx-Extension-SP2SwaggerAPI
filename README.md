## SP2SwaggerAPI

This is a SPfx Extension for sendging Modern page structure to a Swagger API server.

### Motivation
MS Flow has a problem extracting Modern Pages to send API, it break some of html structure expected by API.

This extension basically exports XML structure as a whole and send to an API server.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO

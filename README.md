## react-tabbed-grouped-list

A SharePoint Framework web part that demonstrates a way to organize and display SharePoint list items using PnPJS and the Pivot and GroupedList Fabric UI components

### Building the code

```bash
git clone the repo
npm i
gulp serve
```

### Install
```bash
gulp build
gulp bundle --ship
gulp package-solution --ship
upload .sppkg to SharePoint app catalog
```
Note: solution currently requires a specific list/columns to be set up as detailed in the blog post: https://spmaestro.com/fabric-ui-react-groupedlists-inside-pivot-tabs/

TODO: Genericize solution to not require specific list

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp build
gulp serve
gulp bundle
gulp package-solution

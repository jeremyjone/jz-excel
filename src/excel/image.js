export class ParseImage {
  constructor(files) {
    this.files = files;

    this.images = {};
    for (const fileKey in files) {
      if (fileKey.indexOf("xl/media/") > -1) {
        const fileNameArr = fileKey.split(".");
        const suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();

        if (
          suffix in
          {
            png: 1,
            jpeg: 1,
            jpg: 1,
            gif: 1,
            bmp: 1
          }
        ) {
          this.images[fileKey] = files[fileKey];
        }
      }
    }
  }

  getImageByName(pathName) {
    if (pathName in this.images) {
      const base64 = this.images[pathName];
      return new Image(pathName, base64);
    }

    return null;
  }
}

class Image {
  constructor(pathName, base64) {
    this.pathName = pathName;
    this.src = base64;
  }
}

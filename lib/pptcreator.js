class Presentation {
    constructor(keepRatio, debug) {
      this.content = {};
      this.slides = [];
      this.newImages = [];
      this.relsPath = "ppt/slides/_rels/";
      this.mediaPath = "ppt/media/";
      this.lastSlideId = null;
      this.isDebug = debug;
      this.keepRatio = keepRatio;
    };
  
    checkIfIsSlide(name) {
  
      let isSlidePattern = /(ppt\/slides\/).*(.xml)(?!.*?rels)/i;
  
      let res = isSlidePattern.exec(name);
  
      if (res != null)
        return true;
  
      return false;
    }
  
    escapeHtml(string) {
  
      if (string == null)
        string = "";
  
      string = string.toString();
  
      return string.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    }

    checkIfHasImage(n, slide) {
      let reg = new RegExp("title=\"Image" + n + "\"");
      let obj = reg.exec(slide);
      return obj != null;
    }
  
    processSlides(pptItem) {
  
      let nslide = "newSlide";
      let slidePath = "ppt/slides/";
      let cmonImgName = "common";
  
      if (pptItem.data == null || pptItem.data.length === 0)
        return;
  
      let lastSlide = this.slides[this.slides.length - 1];
  
      if (pptItem.imgData != null)
        for (let x = 0; x < pptItem.imgData.length; x++) {
          pptItem.imgData[x].cid = cmonImgName + x;
        }
  
      this.slides.splice(this.slides.length - 1, 1);
  
      let lastSlideRels = lastSlide.replace(slidePath, this.relsPath) + ".rels";
  
      let cloneSlide = {
        slide: this.content[lastSlide],
        rels: this.content[lastSlideRels]
      }
  
      delete this.content[lastSlide];
      delete this.content[lastSlideRels];
      this.removeSlideFromContent(lastSlide);
  
      let imgCount = 1;
  
      while (true) {
  
        if (!this.checkIfHasImage(imgCount, cloneSlide.slide)) {
          if (imgCount > 1)
            imgCount--;
          break;
        }
        imgCount++;
      }
      let x = 0;
      while (x < pptItem.data.length) {
        let obj = {
          slideContent: cloneSlide.slide,
          referenceContent: cloneSlide.rels
        }
  
        let nSlidePath = slidePath + nslide + x + ".xml";
        let nSlideRelsPath = this.relsPath + nslide + x + ".xml.rels";
  
        for (let imgx = 1; imgx <= imgCount; imgx++) {
  
          for (let prop in pptItem.data[x]) {
            if (pptItem.data[x].hasOwnProperty(prop) && prop != "imgData") {
              let propName = prop;
              if (imgCount > 1)
                propName = prop + imgx;
  
              obj.slideContent = obj.slideContent.replace(new RegExp(propName, 'g'), this.escapeHtml(pptItem.data[x][prop]));
            }
          }
  
          let i = 0;
          for (let i = 0; i < pptItem.data[x].imgData.length; i++) {
  
            if (pptItem.data[x].imgData[i] != null)
              this.replaceImages(obj, pptItem.data[x].imgData[i], x + "_" + i, imgCount > 1 ? imgx : null);
          }
  
          x++;
  
          if (x >= pptItem.data.length)
            break;
  
        }
  
        if (pptItem.imgData != null)
          //Logo reaproveita as imagens
          for (let z = 0; z < pptItem.imgData.length; z++) {
            this.replaceImages(obj, pptItem.imgData[z], pptItem.imgData[z].cid);
          }
  
        this.content[nSlideRelsPath] = obj.referenceContent;
        this.content[nSlidePath] = obj.slideContent;
        this.addSlideToContent(nSlidePath);
        this.slides.push(nSlidePath);
      }
  
      this.configureImageTypesInContent();
  
      this.content["docProps/app.xml"] = this.content["docProps/app.xml"].replace(/<Slides>\d<\/Slides>/, "<Slides>" + this.slides.length + "</Slides>");
    }
  
    configureImageTypesInContent() {
      let imageList = [{
          ext: "png",
          value: "<Default Extension=\"png\" ContentType=\"image/png\"/>"
        },
        {
          ext: "jpg",
          value: "<Default Extension=\"jpg\" ContentType=\"image/png\"/>"
        },
        {
          ext: "jpeg",
          value: "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>"
        }
      ];
  
      for (let x = 0; x < imageList.length; x++) {
        if (this.content["[Content_Types].xml"].indexOf("<Default Extension=\"" + imageList[x].ext + "\"") == -1) {
  
          let index = this.content["[Content_Types].xml"].indexOf("<Default ");
          this.content["[Content_Types].xml"] = this.content["[Content_Types].xml"].substring(0, index) + imageList[x].value + this.content["[Content_Types].xml"].substring(index);
        }
      }
    }
  
  
    removeSlideFromContent(slideName) {
      let contentPart = "<Override PartName=\"/{0}\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>";
  
      this.content["[Content_Types].xml"] = this.content["[Content_Types].xml"].replace(contentPart.replace("{0}", slideName), "");
  
      let rel = new RegExp("<Relationship (?:(?!.\/>).)*(?=Target=\"{0}\").*?\/>".replace("{0}", slideName.replace("ppt/", "")));
  
      let relPath = "ppt/_rels/presentation.xml.rels";
      let presPath = "ppt/presentation.xml";
  
      let res = rel.exec(this.content[relPath]);
  
      if (res != null) {
  
        let id = /(Id=\")(\w+)\"/.exec(res[0]);
  
        let idPattern = new RegExp("(<p:sldId) (?:(?!.\/>).)*r:id=\"{0}\".*?\/>".replace("{0}", id[2]));
  
        this.content[presPath] = this.content[presPath].replace(idPattern, "");
      }
  
      this.content[relPath] = this.content[relPath].replace(rel, "");
  
    }
  
    addSlideToContent(slideName) {
      let contentPart = "<Override PartName=\"/{0}\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>";
      let relPath = "ppt/_rels/presentation.xml.rels";
      let presPath = "ppt/presentation.xml";
      let slid = "<p:sldId id=\"{0}\" r:id=\"{1}\" />";
  
      let contentWithName = contentPart.replace("{0}", slideName);
  
      let xml = this.content["[Content_Types].xml"];
  
      let index = xml.indexOf("<Override");
  
      this.content["[Content_Types].xml"] = xml.substring(0, index) + contentWithName + xml.substring(index);
  
      // Presentation
  
      let relationship = "<Relationship Id=\"{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"{1}\"/>";
      let nid = "nid" + this.slides.length;
  
      let newRelation = relationship.replace("{0}", nid).replace("{1}", slideName.replace("ppt/", ""));
      index = this.content[relPath].indexOf("<Relationship ");
  
      this.content[relPath] = this.content[relPath].substring(0, index) + newRelation + this.content[relPath].substring(index);
  
      index = this.content[presPath].indexOf("</p:sldIdLst>");
  
      if (this.lastSlideId == null) {
        // Pega o ultimo ID
  
        let slidReg = /(<p:sldId) (?:(?!.\/>).)*(id=\")(\d+)\".*?\/>/g;
  
        let resSlid = slidReg.exec(this.content[presPath]);
  
        if (resSlid == null) {
          this.lastSlideId = 256;
          slid = slid.replace("{0}", this.lastSlideId).replace("{1}", nid);
        } else {
          while (resSlid = slidReg.exec(this.content[presPath])) {
            this.lastSlideId = parseInt(resSlid[resSlid.length - 1]) + 1;
          }
          slid = slid.replace("{0}", ++this.lastSlideId).replace("{1}", nid);
        }
      } else {
        slid = slid.replace("{0}", ++this.lastSlideId).replace("{1}", nid);
      }
  
      this.content[presPath] = this.content[presPath].substring(0, index) + slid + this.content[presPath].substring(index);
    }
  
    replaceImages(obj, image, cid, n) {
  
      let imageName = image.name;
  
      if (n != null) {
        imageName += n;
      }
  
      let picPattern = "<p:pic>(?:(?!.<\/p:pic>).)*(?={0}).*?<\/p:pic>".replace("{0}", imageName);
      let psizePattern = "<p:pic>(?:(?!.<\/p:pic>).)+(?={0}).+?(<a:ext cx=\"\\d*\" cy=\"\\d*\"\/>).*?<\/p:pic>".replace("{0}", imageName);
     // let regPic = new RegExp(picPattern);
      let regPSize = new RegExp(psizePattern);
     // let result = regPic.exec(obj.slideContent);
      let result = regPSize.exec(obj.slideContent);
      let ids = [];
      let nid = "new" + cid;
  
      let ext;
  
      if (image.Type == 'jpg')
        ext = ".jpg";
      else
        ext = ".png";
  
      let newImage = "newImage" + cid + ext;
  
      this.content[this.mediaPath + newImage] = image.data;
  
      if (result != null) {
          let res = /(?:embed=").*?"/.exec(result[0]);
          if (res != null) {
            let oldId = res[0].replace("embed=\"", "").replace("\"", '');
          
          let cx = parseInt(/(?:cx="(\d*)")/.exec(result[1])[1]);
          let cy = parseInt(/(?:cy="(\d*)")/.exec(result[1])[1]);
          let ratio = 0;
          let newsize;
          let newContent = result[0];
  
          if (this.keepRatio) {
            if (image.Width > image.Height) {          
              ratio = image.Width / image.Height;
              cy = Math.round(cy / ratio);
              newsize = result[1].replace(/(cy=")+\d*"/,"cy=\""+cy+"\"");          
    
            } else if (image.Height > image.Width) {
              ratio = image.Height / image.Width;
              cx = Math.round(cx / ratio);
              newsize = result[1].replace(/(cx=")+\d*"/,"cx=\""+cx+"\"");
            }

            if (newsize) {
              newContent = result[0].replace(result[1],newsize);
            }
          }
            obj.slideContent = obj.slideContent.replace(result[0], newContent.replace(res[0], res[0].replace(oldId, nid)));
          }
  
        let imgRelationship = "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/{0}\" Id=\"{1}\" />";
  
        imgRelationship = imgRelationship.replace("{0}", newImage).replace("{1}", nid);
  
        let index = obj.referenceContent.indexOf("</Relationships>");
  
        obj.referenceContent = obj.referenceContent.substring(0, index) + imgRelationship + obj.referenceContent.substring(index);
  
      }
    }
  
    load(data) {
  
      let promise = new Promise(resolve => {
      let obj = this;

      const convertSlidesToObjs = (key, zip) => {
          let _promise = new Promise(res => {
              let ext = key.substr(key.lastIndexOf('.'));
    
            //Checa se é um slide padrão
            if (obj.checkIfIsSlide(key))
                this.slides.push(key);

            if (ext == '.xml' || ext == '.rels') {
              
              zip.file(key).async("string").then(function(xml) {    
                obj.content[key] = xml;
                res();
              })
    
            } else {
    
            zip.file(key).async("base64").then(function(img) {
              obj.content[key] = img;
              res();
            })
          }
        });

        return _promise;
      };
      
  
      JSZip.loadAsync(data).then(function(zip) {        
        let promises = Object.keys(zip.files).map(obj => convertSlidesToObjs(obj, zip));
        
        Promise.all(promises).then(() =>
         { 
           if (obj.isDebug)
             console.log("Loaded");
           resolve();
          });
      })
    });

    return promise;
    };
  
    loadImage(obj)  {                
      var promise = new Promise( resolve=> {
          var image = new Image();
          var canvas = document.createElement("canvas"),
          canvasContext = canvas.getContext("2d");
          image.crossOrigin = "Anonymous";
          image.onload = function() {

              canvas.width = image.width;
              canvas.height = image.height;
              canvasContext.drawImage(image, 0, 0, image.width, image.height);
              var type = "";

              if (obj.path.indexOf(".png") != -1)
                type = "image/png";
              else
                type = "image/jpeg";

              var data = canvas.toDataURL(type).split(',')[1];
              obj.data = data;
              resolve();
          }

            image.src = obj.path;                        
          });

          return promise;
    }

    loadAndProcess(objs, template) {
      let promise = new Promise(resolve =>{ 
        if (objs.data == null || objs.data.length == 0) {
          throw "No data in object";
        }

        let imgPromises = [];
        // check for common image
        if (objs.imgData != null && objs.imgData.length > 0) {
            imgPromises.push(objs.imgData.map(item => this.loadImage(item)));
        }
        // check for images path
        objs.data.forEach(element => {
            if (element.imgData != null && element.imgData.length > 0) {
                imgPromises.push(element.imgData.map(item => this.loadImage(item)));
            }
        });

        Promise.all(imgPromises).then(() =>
         this.loadTemplate(template).then(tempBlob => {
          this.load(tempBlob).then(() => {
            this.processSlides(objs);
            resolve();
          })
        }));
      });
    return promise;
    }
  
    loadTemplate(path) {
                    
      let promise = new Promise(function(resolve, reject) {
         fetch(path)
             .then(res => res.blob()) // Gets the response and returns it as a blob
             .then(blob => {
             resolve(blob);
             });
         })
         return promise;
  }

  downloadContent(content) {
      var blob = new Blob([content]);
      var link = document.createElement('a');
      link.href = window.URL.createObjectURL(blob);
      link.download = "PhotoBook.pptx"
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
  }

  downloadBuffer() {
    this.toBuffer().then(content => this.downloadContent(content));
  }
  
    toBuffer() {
      let zip2 = new JSZip();
      let content = this.content;
      for (let key in content) {
        if (content.hasOwnProperty(key)) {
          let ext = key.substr(key.lastIndexOf('.'));
          if (ext == '.xml' || ext == '.rels') {            
            zip2.file(key, content[key]);
          } else {
            zip2.file(key, content[key], {
              base64: true
            });
          }
        }
      }      
      zip2.file("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Words>0</Words><Application>Microsoft Macintosh PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>2</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr></vt:vector></TitlesOfParts><Company>Proven, Inc.</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>')
      return zip2.generateAsync({
        type: "blob"
      });
  
    };
  }
  
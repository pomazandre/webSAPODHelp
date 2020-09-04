const Decompress = require('decompress');
var config = require('../config.json');
var Fs = require('fs');
var DOMParser = require('xmldom').DOMParser;

function HelpWriter() {
    this.contentListItems = [];
    this.contentHyperLinks = [];
    this.Observables = [];
    this.hyperLinkIndexes = [];
    this.curChapters = [];

    this.curChapters = [];
    this.Content = undefined; // по параграфам
    this.Doc = undefined; // по параграфам
}

HelpWriter.prototype.perform = function () {
    self = this;
    Decompress(config.source_doc, 'tmp').then(files => {
        self.parse(files);
    }
    );
}


HelpWriter.prototype.init = function (Data) {
    this.Doc = new DOMParser().parseFromString(Data.toString(), 'utf-8', 'application/xml');
    this.Content = this.Doc.getElementsByTagName('w:p'); // по параграфам
}


HelpWriter.prototype.parseContentList = function () {
    console.log('парсинг оглавления ...');
    this.Content = this.Doc.getElementsByTagName('w:p'); // по параграфам
    ParaLength = this.Content.length;
    paraMaxIndex = 0;
    for (var i = 0; i < ParaLength; i++) {

        for (var j = 0; j < this.Content[i].childNodes.length; j++) { // либо ищем тег  instrText либо тег w:hyperlink
            const _Node = this.Content[i].childNodes[j];
            if (_Node.nodeName === 'w:hyperlink') {
                _hyperLink = _Node.getAttribute("w:anchor");
                curListCounter = 0;
                contentListItem = '';
                for (var k = 0; k < _Node.childNodes.length; k++) {
                    nodeChild = _Node.childNodes[k];
                    if (nodeChild.nodeName === "w:r") {
                        curListItem = '';
                        for (var l = 0; l < nodeChild.childNodes.length; l++) {
                            nodeChildChild = nodeChild.childNodes[l];
                            if (nodeChildChild.nodeName === 'w:t') {
                                if (nodeChildChild.childNodes.length === 1) {
                                    if (nodeChildChild.childNodes[0].nodeValue !== null) {
                                        curListItem = curListItem + nodeChildChild.childNodes[0].nodeValue;
                                        curListCounter++;
                                    }
                                }
                            }
                            if (curListCounter <= 2) {
                                contentListItem = contentListItem + curListItem;
                            }
                        }
                    }
                }
                if (contentListItem !== '') {
                    this.contentListItems.push(contentListItem);
                    this.contentHyperLinks.push(_hyperLink);
                    paraMaxIndex = i;
                }
            }
        }

    }
    console.log('парсинг оглавления завершен');
}

HelpWriter.prototype.saveIndexHtml = function () {
    console.log('сохранение главной страницы...');
    var stream = Fs.createWriteStream(config.out_folder + 'index.html');
    stream.once('open', function (fd) {
        var i;
        stream.write("<!DOCTYPE html>");
        stream.write("<html>");
        stream.write("<head>");
        stream.write("<meta charset=\"utf-8\" />");
        stream.write("<title>HTML5</title>");
        stream.write("</head>");
        stream.write("<body>");
        for (var i = 0; i < self.contentListItems.length; i++) {
            var _nameChapter = getNameChapter(self.contentListItems[i]);
            stream.write("<p><a href=\"" + _nameChapter + "\">" + self.contentListItems[i] + "</a></p>")
        }
        stream.write("</body>");
        stream.write("</html>");
        stream.end();
        console.log('главная страницы сохранена');
    });
}

HelpWriter.prototype.parseContentListIndexes = function () {
    console.log('анализ документа и запись индексов подразделов в массив...');
    for (var i = paraMaxIndex; i < ParaLength - 1; i++) {
        for (var j = 0; j < this.Content[i].childNodes.length; j++) {
            const _Node = this.Content[i].childNodes[j];
            if (_Node.nodeName === 'w:bookmarkStart') {
                _hyperLink = _Node.getAttribute("w:name");
                if (this.contentHyperLinks.indexOf(_hyperLink) >= 0) {
                    this.hyperLinkIndexes.push(i);
                }
            }
        }
    }
}


HelpWriter.prototype.parse = function () {
    var ParaLength, contentListItem, curListItem, _hyperLink, paraMaxIndex;
    var self = this;

    Fs.readFile('./assets/tmp/word/document.xml', function (Error, Data) {
        if (Error) {
            throw Error;
        }
        self.init(Data);
        self.parseContentList();
        self.saveIndexHtml();
        self.parseContentListIndexes();
        self.parseChapters();

    })
}

HelpWriter.prototype.parseChapters = function () {
    console.log('парсинг подразделов ...');
    var curChapter = "";

    for (var i = 0; i < this.hyperLinkIndexes.length; i++) { // идем уже по готовым индексам параграфов
        var startIndex = this.hyperLinkIndexes[i];
        var finalIndex = 0;
        if (i === self.hyperLinkIndexes.length - 1) {
            finalIndex = ParaLength - 1;
        } else {
            finalIndex = self.hyperLinkIndexes[i + 1];;
        }
        console.log(startIndex);
        console.log(finalIndex);
        curChapter = "";
        var curParag;
        for (var j = startIndex; j <= finalIndex; j++) { // цикл по параграфам
            const _Node = self.Content[j];
            curParag = "";
            for (var k = 0; k < _Node.childNodes.length; k++) { // цикл по подузлам
                var curStyleStart = "";
                var curStyleFinish = "";
                const _subNode = _Node.childNodes[k];
                if (_subNode.nodeName === 'w:r') {
                    for (var l = 0; l < _subNode.childNodes.length; l++) {
                        const _subSubNode = _subNode.childNodes[l];
                        if (_subSubNode.nodeName === '<w:rPr>') {
                            for (var s = 0; s < _subSubNode.childNodes.length; s++) {
                                var subSubSubNode = _subSubSubNode.childNodes[s];
                                if (_subSubSubNode.nodeName === 'w:i') {
                                    curStyleStart = '<span style="font-style: italic;">';
                                    curStyleFinish = '</span">';
                                }
                                if (_subSubSubNode.nodeName === 'w:b') {
                                    curStyleStart = '<span style="font-weight: 700;">';
                                    curStyleFinish = '</span">';
                                }
                            }
                        }
                        if (_subSubNode.nodeName === 'w:t') {
                            if (_subSubNode.childNodes.length === 1) {
                                if (_subSubNode.childNodes[0].nodeValue !== null) {
                                    curParag = curParag + curStyleStart + _subSubNode.childNodes[0].nodeValue + curStyleFinish;
                                }
                            }
                        }
                    }
                }
            }
        }
        if (curParag.trim() !== '') {
            if (j === startIndex) {
                curChapter = curChapter + '<h1 align="center">' + curParag + '</h2>';
            } else {
                curChapter = curChapter + '<p align="justify">' + curParag + '</p>';
            }
        }
        var _nameChapter = config.out_folder + getNameChapter(self.contentListItems[i]);
        Fs.writeFileSync(_nameChapter, curChapter);
    }
    console.log('парсинг пдоразделов завершен');
}

function getNameChapter(name) {
    var i;
    var _Pos;
    _Pos = 0;
    for (i = 0; i < name.length; i++) {
        if (name.charAt(i) === '.') {
            _Pos = i;
        }
    }
    return name.substr(0, _Pos) + '.html';
}

module.exports = new HelpWriter();
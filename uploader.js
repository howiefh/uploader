/**
 * 移动端压缩上传图片插件
 * 
 * Zepto 默认版本不能使用， 需要加别的模块
 * @author Feng Hao
 */
(function($, lrz) {

    var Uploader = (function() {

        /**
         * @param element
         * @param options
         * @constructor
         */
        function Uploader(element, options) {
            this.$element = $(element);

            this.settings = $.extend({}, $.fn.Uploader.defaults, options);

            this.browseBtn = $(this.settings.browseButton);
            this.submitBtn = $(this.settings.submitButton);
            this.files = { length: 0 };
            this.uploading = false;
            this.preview = 0;

            this.init();
        }

        Uploader.prototype = {
            init: function() {
                var self = this;
                self.browseBtn.on('change', function(e) {
                    var existFiles = self.files.length;
                    for (var i = 0, file;
                        (file = this.files[i]) && (i + existFiles) < self.settings.maxFiles; i++) {
                        /* 超过最大文件数，隐藏上传按钮 */
                        if (i + existFiles == self.settings.maxFiles - 1) {
                            $('#browse-item').hide();
                        }
                        self.addFile(file);
                    }
                });
                self.submitBtn.on('click', function(e) {
                    if (self.preview > 0) {
                        self.showMsg('正在加载图片, 请您耐心等待');
                        setTimeout(function () {
                            $('.fix-msg').fadeOut();
                        }, 2000);
                        return;
                    }
                    if (!self.uploading) {
                        self.uploading = true;
                        self.showMsg('正在提交中...');
                        self.beforeUpload();
                    } else {
                        self.showMsg('正在提交中, 请您耐心等待');
                    }
                });
            },

            addFile: function(file) {
                var self = this;
                var fileId = _generateUUID();
                self.preview++;
                $('#browse-item').before('<li class="file-item loading" id="' + fileId + '"><span></span></li>');
                lrz(file)
                    .then(function(rst) {
                        var imgBase = rst.base64;
                        if (imgBase.indexOf(";") < 0) {
                            imgBase = imgBase.replace("data:", "data:image/jpeg;")
                        }
                        if (imgBase.indexOf("data:;") > -1) {
                            imgBase = imgBase.replace("data:;", "data:image/jpeg;")
                        }

                        self.files[fileId] = {
                            id: fileId,
                            name: file.name,
                            base64: imgBase,
                            file: rst.file,
                            size: rst.fileLen
                        };
                        self.files.length++;
                        
                        $('#' + fileId).removeClass('loading').html('<a class="delete-img" id="delete-' + fileId + '" href="javascript:;"></a><img class="img img-item" src="' + imgBase + '">');
                        $('#delete-' + fileId).on('click', function() {
                            self.deleteFile(fileId);
                        });
                        self.preview--;
                    })
                    .catch(function(err) {
                        console.info(err)
                    });
            },

            deleteFile: function(fileId) {
                delete this.files[fileId];
                this.files.length--;
                $('#' + fileId).remove();
                if (this.files.length < this.settings.maxFiles) {
                    $('#browse-item').show();
                }
            },

            beforeUpload: function() {
                var self = this;
                setTimeout(function() {
                    var params, 
                    fd = new _FormDataShim();

                    for (var id in self.files) {
                        if (id == 'length') {
                            continue;
                        }
                        fd.append('file', self.files[id].file, self.files[id].name);
                    }

                    try {
                        params = _isFunction(self.settings.multipartParams) ? self.settings.multipartParams() : self.settings.multipartParams;
                    } catch(e) {
                        if (e.name == 'TypeError') {
                            self.uploadError(e.message);
                        } else {
                            self.uploadError('请求参数有误');
                        }
                        return;
                    }

                    for (var name in params) {
                        fd.append(name, params[name]);
                    }

                    $.ajax({
                        url: self.settings.url,
                        type: 'post',
                        data: fd,
                        processData: false,
                        contentType: 'multipart/form-data; boundary=' + fd.boundary,
                        success: function(result) {
                            self.uploadComplete(result);
                        },
                        error: function(XMLHttpRequest, textStatus, errorThrown) {
                            self.uploadError("系统繁忙, 请稍后重试");
                        },
                    });

                }, 1000);

            },

            uploadComplete: function (result) {
                var self = this;
                if (result.code == 200) {
                    this.showMsg('感谢您的反馈, 正在为您跳转链接');
                    setTimeout(function () {
                        window.location.replace(self.settings.returnUrl);
                    }, 2000);
                } else if (result.code == 302) {
                    window.location.href = result.url;
                } else if (result.code == 401) {
                    this.uploadError(result.msg);
                } else {
                    this.uploadError("系统繁忙, 请稍后重试");
                }
            },

            uploadError: function (txt1) {
                var self = this;
                self.showMsg(txt1);
                setTimeout(function () {
                    $('.fix-msg').fadeOut();
                    self.uploading = false;
                }, 2000);
            },

            showMsg: function (txt1) {
                $('#msgTxt_1').html(txt1);
                var msg = $('.fix-msg');
                if(msg.css("display") == "none"){
                    msg.fadeIn();
                }
            }
        };

        return Uploader;

    })();

    function _generateUUID () {
        var d = new Date().getTime();
        var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            var r = (d + Math.random() * 16) % 16 | 0;
            d = Math.floor(d / 16);
            return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
        });
        return uuid;
    };

    var blobConstruct = !!(function () {
        try { return new Blob(); } catch (e) {}
    })(),
    XBlob = blobConstruct ? window.Blob : function (parts, opts) {
        var bb = new (window.BlobBuilder || window.WebKitBlobBuilder || window.MSBlobBuilder);
        parts.forEach(function (p) {
            bb.append(p);
        });
 
        return bb.getBlob(opts ? opts.type : undefined);
    };

    function _FormDataShim () {
        // Store a reference to this
        var o = this,
            parts = [],// Data to be sent
            boundary = Array(5).join('-') + (+new Date() * (1e16*Math.random())).toString(32),
            oldSend = XMLHttpRequest.prototype.send;

        this.boundary = boundary;
     
        this.append = function (name, value, filename) {
            parts.push('--' + boundary + '\r\nContent-Disposition: form-data; name="' + name + '"');
     
            if (value instanceof Blob) {
                parts.push('; filename="'+ (filename || 'blob') +'"\r\nContent-Type: ' + value.type + '\r\n\r\n');
                parts.push(value);
            } else {
                parts.push('\r\n\r\n' + value);
            }
            parts.push('\r\n');
        };
     
        // Override XHR send()
        XMLHttpRequest.prototype.send = function (val) {
            var fr,
                data,
                oXHR = this;
     
            if (val === o) {
                //注意不能漏最后的\r\n ,否则有可能服务器解析不到参数.
                parts.push('--' + boundary + '--\r\n');
                data = new XBlob(parts);
                fr = new FileReader();
                fr.onload = function () { 
                    oldSend.call(oXHR, fr.result); 
                };
                fr.onerror = function (err) { throw err; };

                this.setRequestHeader('Content-Type', 'multipart/form-data; boundary=' + boundary);
                fr.readAsArrayBuffer(data);
     
                XMLHttpRequest.prototype.send = oldSend;
            } else {
                oldSend.call(this, val);
            }
        };
    }

    function _isFunction(fn) {
        return Object.prototype.toString.call(fn)=== '[object Function]';
    }


    $.fn.Uploader = function(options) {
        return this.each(function() {
            var self = $(this),
                instance = $.fn.Uploader.lookup[self.data('uploader')];
            if (!instance) {
                //zepto的data方法只能保存字符串，所以用此方法解决一下
                $.fn.Uploader.lookup[++$.fn.Uploader.lookup.i] = new Uploader(this, options);
                self.data('plugin', $.fn.Uploader.lookup.i);
                instance = $.fn.Uploader.lookup[self.data('uploader')];
            }

            if (typeof options === 'string') instance[options]();
        });
    };
    $.fn.Uploader.lookup = { i: 0 };

    $.fn.Uploader.defaults = {
        browseButton: '#browse',
        submitButton: '#submit',
        url: '/file/upload',
        returnUrl: '',
        maxFiles: 5,
        multipartParams: null
    };

    $(function() {
        // return new Uploader($('[data-uploader]'));
    });
})(window.Zepto || window.jQuery, lrz);

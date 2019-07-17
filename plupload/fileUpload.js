/**
 * @author Feng Hao
 */
(function($){
    'use strict';

    var defaults = {
        runtimes: 'html5,flash,silverlight,html4',
        browse_button : 'uploadFile', // you can pass in id...
        // Flash settings
        flash_swf_url: '/plugins/plupload/Moxie.swf',
        // Silverlight settings
        silverlight_xap_url: '/plugins/plupload/Moxie.xap',

        autostart: true,

        multi_selection: false,

        fileNameRule: null,

        fileNameMsg: null,

        fileNameReplaceOld: [],

        fileNameReplaceNew: [],

        fileNameMaxLength: 0,

        successFn: function (){},

        errorFn: function (){},

        init: {
            PostInit: function(up) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                browseBtn.attr('title', '');
            },

            FilesAdded: function(up, files) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                var err = false;

                plupload.each(files, function(file) {
                    if (opt.fileNameReplaceOld) {
                        for (var i = 0; i < opt.fileNameReplaceOld.length; i++) {
                            file.name = file.name.replace(opt.fileNameReplaceOld[i], opt.fileNameReplaceNew[i])
                        }
                    }
                    if (opt.fileNameMaxLength > 0 && opt.fileNameMaxLength < file.name.length) {
                        alert(file.name + ': ' + '文件名长度超出最大长度 ' + opt.fileNameMaxLength + '，请修改文件名后重新上传');
                        err = true;
                        return false;
                    }
                    if (opt.fileNameRule && opt.fileNameRule instanceof RegExp && !opt.fileNameRule.test(file.name)) {
                        alert(file.name + ': ' + (opt.fileNameMsg || '文件名不规范，请修改文件名后重新上传'));
                        err = true;
                        return false;
                    }
                });

                if (err) {
                    // 清空文件
                    up.splice(0, up.files.length);
                    return;
                }

                plupload.each(files, function(file) {
                    browseBtn.attr('disabled', true).attr('title', file.name + ' (' + plupload.formatSize(file.size) + ') ').attr('data-loading', '');
                });
                //自动上传
                if (opt.autostart) {
                    setTimeout(function () {
                        up.start();
                    }, 10);
                }
            },

            UploadProgress: function(up, file) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                var percent = file.percent;
                if (percent == 100) {
                    percent = 99;
                }
                browseBtn.find('.text').html(percent + '%');
                browseBtn.find('.process-bar').css('width', percent + '%');
            },

            FileUploaded: function(up, file, info) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                var response = JSON.parse(info.response);
                // alert(response.msg);
                var textEle = browseBtn.find('.text');
                textEle.html(textEle.data('text'));
                browseBtn.attr('disabled', false).removeAttr('data-loading');
                if ($.isFunction(opt.successFn)) {
                    opt.successFn(response, file);
                }
            },

            Error: function(up, err) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                alert("上传失败 " + err.code + ": " + err.message);
                if ($.isFunction(opt.errorFn)) {
                    opt.errorFn(err);
                }
                var textEle = browseBtn.find('.text');
                textEle.html(textEle.data('text'));
                browseBtn.attr('disabled', false).removeAttr('data-loading');
            }
        }
    };

    var FileUpload = function(element, options){
        var self = this;
        self.opts = $.extend(true, {}, defaults, options);
        var browseButton = self.opts.browse_button, btn;
        if (Array.isArray(browseButton)) {
            // 数组中只能有一个生效可见
            for (var i = 0; i < browseButton.length; i++) {
                btn = $('#' + browseButton[i]);
                if (btn.length > 0) {
                    self.opts.browse_button = browseButton[i];
                    break;
                }
            }
        } else {
            btn = $('#' + browseButton);
        }

        if (btn.length > 0) {
            var processBar = btn.find('.progress-bar');
            // 没有 processBar 子元素 且 只含文本子节点
            if (processBar.length === 0 && btn.children().length === 0) {
                var text = btn.text() && btn.text().trim();
                btn.html('<span class="text" data-text="' + text + '">' + btn.html() + '</span><div class="process-bar"></div>')
            }
            var uploader = new plupload.Uploader(self.opts);
            uploader.init();
            self.uploader = uploader;
        } else {
            console.error('Can not find element browse_button: ' + self.opts.browse_button);
        }

    };

    $.fn.FileUpload = function(options){
        return new FileUpload(this, options);
    };

})(jQuery);


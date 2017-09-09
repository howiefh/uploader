/**
 * @author: Feng Hao  
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

        successFn: null,

        errorFn: null,

        init: {
            PostInit: function(up) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                browseBtn.attr('title', '');
            },

            FilesAdded: function(up, files) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
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
                if ($.isFunction(opt.successFn)) {
                    opt.successFn(response);
                }
                var textEle = browseBtn.find('.text');
                textEle.html(textEle.data('text'));
                browseBtn.attr('disabled', false).removeAttr('data-loading');
            },

            Error: function(up, err) {
                var opt = up.getOption();
                var browseBtn = $(opt.browse_button);
                alert("Error #" + err.code + ": " + err.message);
                if ($.isFunction(opt.errorFn)) {
                    opt.errorFn();
                }
                var textEle = browseBtn.find('.text');
                textEle.html(textEle.data('text'));
                browseBtn.attr('disabled', false).removeAttr('data-loading');
            }
        }
    };

    var FileUploader = function(element, options){
        var self = this;
        self.opts = $.extend(true, {}, defaults, options);
        var btnId = self.opts.browse_button;
        var btn = $('#' + btnId);
        var processBar = btn.find('.progress-bar');
        if (processBar.length == 0) {
            btn.html('<span class="text" data-text="' + btn.html() + '">' + btn.html() + '</span><div class="process-bar"></div>')
        }
        var uploader = new plupload.Uploader(self.opts);
        uploader.init();
        self.uploader = uploader;
    };

    $.fn.fileUploader = function(options){
        return new FileUploader(this, options);
    };

})(jQuery);


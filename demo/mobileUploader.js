FastClick && FastClick.attach(document.body)

$('.upload-area').Uploader({
    url: '',
    returnUrl: '',
    multipartParams: function(){
        var surveyContent = $.trim($('#surveyContent').val()),
        mobileNo = $.trim($('#mobileNo').val()),
        phone = /^(130|131|132|133|145|134|135|136|137|138|139|147|150|151|152|153|155|156|157|158|159|170|173|176|177|178|180|181|182|183|184|185|186|187|188|189)[0-9]{8}$/,
        content = /^.{1,150}$/;

        if (surveyContent == '') {
            throw new TypeError('请输入您的问题或建议');
        }

        if (!content.test(surveyContent)) {
            throw new TypeError('最多可输入150个字符');
        }

        if (!phone.test(mobileNo)) {
            throw new TypeError('请输入合法手机号');
        }

        return {
            surveyContent: surveyContent,
            mobileNo: mobileNo
        }
    }
});
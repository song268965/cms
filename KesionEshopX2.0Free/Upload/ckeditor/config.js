/**
 * @license Copyright (c) 2003-2013, CKSource - Frederico Knabben. All rights reserved.
 * For licensing, see LICENSE.md or http://ckeditor.com/license
 */

CKEDITOR.editorConfig = function( config ) {
	// Define changes to default configuration here. For example:
	// config.language = 'fr';
	// config.uiColor = '#AADC6E';
	
	config.language = 'zh-cn';//中文
    config.font_names = '宋体;楷体_GB2312;新宋体;黑体;隶书;幼圆;微软雅黑;Arial;Comic Sans MS;Courier New;Tahoma;Times New Roman;Verdana';	
	config.uiColor = '#F1F1F1';
	config.enterMode = CKEDITOR.ENTER_BR;
    config.shiftEnterMode = CKEDITOR.ENTER_P;
	config.pasteFromWordIgnoreFontFace = false;
	config.pasteFromWordKeepsStructure = true;
	config.pasteFromWordRemoveStyle = false;
	config.forcePasteAsPlainText =false ;

//改变大小的最大高度
    config.resize_maxHeight = 3000;
    //改变大小的最大宽度
    config.resize_maxWidth = 3000;
    //改变大小的最小高度
    config.resize_minHeight = 250;
    //改变大小的最小宽度
    config.resize_minWidth = 650;
	
	

	config.toolbar_Default =
    [
        ['Source','-','Preview','-','Templates'],
        ['Cut','Copy','Paste','PasteText','PasteFromWord','-','Print', 'SpellChecker', 'Scayt'],
        ['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
        ['Form', 'Checkbox', 'Radio', 'TextField', 'Textarea', 'Select', 'Button', 'ImageButton', 'HiddenField'],
        '/',
        ['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
        ['NumberedList','BulletedList','-','Outdent','Indent','Blockquote','CreateDiv'],
        ['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
        ['Link','Unlink','Anchor'],
          ['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar','PageBreak'],
        '/',
        ['Styles','Format','Font','FontSize'],
        ['TextColor','BGColor'],
        ['Maximize', 'ShowBlocks','-','About']
    ];
	config.toolbar_NewsTool =
    [
        ['Source','-','Preview','-','Templates'],
        ['Cut','Copy','Paste','PasteText','PasteFromWord','-','Print', 'SpellChecker', 'Scayt'],
        ['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
        ['Link','Unlink','Anchor'],
        '/',
        ['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
        ['NumberedList','BulletedList','-','Outdent','Indent'],
        ['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
          ['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar'],
        '/',
        ['Styles','Format','Font','FontSize'],
        ['TextColor','BGColor'],
        ['Maximize', '-','About']
    ];
	config.toolbar_Basic =
    [
        ['Bold', 'Italic','Underline','TextColor','BGColor', '-','FontSize', 'NumberedList', 'BulletedList', '-', 'Link', '-','Image','Flash','Table']
    ];
	config.toolbar_Simple =
    [
        ['Preview','-','Templates'],
        ['Cut','Copy','Paste','PasteText','PasteFromWord','-','Print', 'SpellChecker', 'Scayt'],
        ['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
        ['Link','Unlink','Anchor'],
        '/',
        ['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
        ['NumberedList','BulletedList','-','Outdent','Indent','Blockquote','CreateDiv'],
        ['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
          ['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar'],
        '/',
        ['Styles','Format','Font','FontSize'],
        ['TextColor','BGColor'],
        ['Maximize', 'ShowBlocks','-','About']
    ];
    
	
};

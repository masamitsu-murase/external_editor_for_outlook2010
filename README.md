
# External Editor for Outlook 2010

This is VBA macro to edit mail content with an external editor in Outlook 2010.

## How to use

### Preparation

1. Launch Outlook 2010.
2. Open VB Editor.  
   Press "Alt" + "F11" to open it.
3. Create a new module and rename it to "ExternalEditor".
4. Copy and paste source code of "external_editor.vbs" to the module.
5. Configure `EDITOR_PATH` and `TEMP_DIR` based on your environment.
6. Open "ThisOutlookSession" and add the following source code.  
   This is not mandatory, but it ensures to release all resources.  
   ```vbnet
   Private Sub Application_Quit()
     ExternalEditor.FinishOpenInExternalEditor
   End Sub
   ```
7. Save files.  
   Sign your macro digitally if you need.

### Run this macro

1. Click "New Mail" to open a new mail.
2. Run "ExternalEditor.OpenInExternalEditor".  
   Your external editor is opened.
3. Edit the mail content and save it with your external editor.
4. When you close the external editor, the content will be copied to the mail editor in Outlook.


## Configurations

You can configure the environment of your external editor with the following constants. You can find them at the beginning of "external_editor.vbs".

* `EDITOR_PATH`  
  Absolute path to your external editor  
  e.g. If you use [xyzzy](https://github.com/xyzzy-022/xyzzy) as your external editor, set `EDITOR_PATH` to `"C:\...\xyzzy.exe"`.
* `TEMP_DIR`  
  Absolute path to the directory in which temporary files are saved  
  This variable must end with "\".
* `REMOVE_TEMP_FILES`  
  If this is set to `True`, this macro removes temporary files after you close the external editor.  
  If this is set to `False`, this macro leave temporary files for fail safe.


## Note and limitations

* If you edit your mail content within the mail editor of Outlook while you edit it within the external editor, the content edited in Outlook editor will be lost when you close the external editor.
* This macro is under development. If you are unlucky, you may face a crash of Outlook...


## History

* 1.0.0  
  Release of the initial version


## License

You can use this software under the MIT License.

Copyright (c) 2016 Masamitsu MURASE

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


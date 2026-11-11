using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using System.IO.Compression;
using System.Collections.Generic;
using System.Windows.Forms;

partial class FormulaHelper {
    dynamic clipboardApp,clipboardDocument;
    bool clipboardPresentation;
    void RestoreClipboardText(string text){if(text!=null)try{Clipboard.SetDataObject(text,true,10,100);}catch(Exception ex){File.AppendAllText(Path.Combine(root,"output","clipboard.log"),DateTime.Now+" "+ex.Message+Environment.NewLine);}}
    const string M="http://schemas.openxmlformats.org/officeDocument/2006/math";
    const string W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string A="http://schemas.openxmlformats.org/drawingml/2006/main";
    const string P="http://schemas.openxmlformats.org/presentationml/2006/main";
    string Omml(Dictionary<string,object> data) {
        string stylesheet=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),"Microsoft Office","root","Office16","MML2OMML.XSL");
        if(!File.Exists(stylesheet))throw new Exception("未找到 Office 的 MathML 转换器 MML2OMML.XSL，请检查桌面版 Office 是否完整安装。");
        var transform=new XslCompiledTransform();transform.Load(stylesheet);
        var output=new StringBuilder();var settings=transform.OutputSettings.Clone();settings.OmitXmlDeclaration=true;
        using(var input=XmlReader.Create((string)data["mathml"],new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit}))
        using(var writer=XmlWriter.Create(output,settings))transform.Transform(input,writer);
        var xml=new XmlDocument();xml.LoadXml(output.ToString());
        var ns=new XmlNamespaceManager(xml.NameTable);ns.AddNamespace("m",M);
        if(((string)data["latex"]).TrimStart().StartsWith(@"\begin{gathered}",StringComparison.Ordinal)){
            // Office does not retain TeX row-gap lengths: set native double row gaps.
            foreach(XmlNode rows in xml.SelectNodes("//m:oMath/m:m|//m:oMath/m:eqArr",ns)){
                string name=rows.LocalName=="m"?"mPr":"eqArrPr";
                XmlNode pr=rows.SelectSingleNode("m:"+name,ns);if(pr==null){pr=xml.CreateElement("m",name,M);rows.PrependChild(pr);}
                var old=pr.SelectSingleNode("m:rSpRule",ns);if(old!=null)pr.RemoveChild(old);
                old=pr.SelectSingleNode("m:rSp",ns);if(old!=null)pr.RemoveChild(old);
                var rule=xml.CreateElement("m","rSpRule",M);rule.SetAttribute("val",M,"2");pr.PrependChild(rule);
            }
        }
        // The Office stylesheet treats MathJax's spacing dot accents as upper limits.
        // Use native math accents so derivative dots stay attached when editing.
        var upperLimits=new List<XmlNode>();foreach(XmlNode node in xml.SelectNodes("//m:limUpp",ns))upperLimits.Add(node);
        foreach(var limit in upperLimits){
            var marker=limit.SelectSingleNode("m:lim",ns);if(marker==null)continue;
            string glyph=marker.InnerText;string combining=glyph=="˙"?"\u0307":glyph=="¨"?"\u0308":null;if(combining==null)continue;
            var accent=xml.CreateElement("m","acc",M);var pr=xml.CreateElement("m","accPr",M);var chr=xml.CreateElement("m","chr",M);chr.SetAttribute("val",M,combining);pr.AppendChild(chr);accent.AppendChild(pr);
            accent.AppendChild(limit.SelectSingleNode("m:e",ns).CloneNode(true));limit.ParentNode.ReplaceChild(accent,limit);
        }
        var eq=xml.SelectSingleNode("//m:oMath",ns);if(eq==null)throw new Exception("公式转换未生成原生数学对象。");File.WriteAllText(Path.Combine(root,"temp","last-omml.xml"),eq.OuterXml);return eq.OuterXml;
    }
    void InsertEditable(string targetName) {
        if(targetName=="Visio"){Insert(false);return;}
        if(targetName.StartsWith("WPS",StringComparison.Ordinal)){
            CopyEditableFormula(targetName);return;
        }
        // Resolve target before rendering. No user document is changed by preview/conversion.
        dynamic app=targetName=="Visio"?Visio():Office(targetName);
        var data=Render();string omml=Omml(data);
        if(targetName=="Word"||targetName=="WPS 文字")InsertNativeWord(app,omml);
        else InsertNativePpt(app,omml,data);
    }
    void ReleaseNativeClipboardSource(){
        if(clipboardDocument!=null)try{if(clipboardPresentation)clipboardDocument.Close();else clipboardDocument.Close(0);}catch{}
        if(clipboardApp!=null)try{if(clipboardPresentation){if(clipboardApp.Presentations.Count==0)clipboardApp.Quit();}else if(clipboardApp.Documents.Count==0)clipboardApp.Quit(0);}catch{}
        clipboardDocument=null;clipboardApp=null;
    }
    void CopyEditableFormula(string targetName){
        if(targetName.StartsWith("WPS",StringComparison.Ordinal)){ReleaseNativeClipboardSource();CopyLatexSource();return;}
        var data=Render();string omml=Omml(data);ReleaseNativeClipboardSource();
        clipboardPresentation=targetName=="PowerPoint"||targetName=="WPS 演示";
        try{
            if(!clipboardPresentation){
                clipboardApp=Activator.CreateInstance(Type.GetTypeFromProgID(targetName=="WPS 文字"?"KWPS.Application":"Word.Application"));if(targetName!="WPS 文字"){clipboardApp.Visible=false;clipboardApp.DisplayAlerts=0;}
                clipboardDocument=clipboardApp.Documents.Open(CreateEquationDocx(omml,null),false,true);
                clipboardDocument.OMaths[1].Range.Copy();
            }else{
                clipboardApp=Activator.CreateInstance(Type.GetTypeFromProgID(targetName=="WPS 演示"?"KWPP.Application":"PowerPoint.Application"));
                string file=NativePptFile(clipboardApp,omml,data);clipboardDocument=clipboardApp.Presentations.Open(file,-1,0,-1);clipboardDocument.Windows[1].WindowState=2;
                clipboardDocument.Slides[1].Shapes[1].Copy();
            }
            System.Threading.Thread.Sleep(200);
            status.Text="排版公式对象已复制。LaTeX 代码文本请使用“复制 LaTeX 源码”。";
        }catch{ReleaseNativeClipboardSource();throw;}
    }
    void PutWordEquation(dynamic range,string omml) {
        string file=CreateEquationDocx(omml,null);range.InsertFile(file);
        if(range.OMaths.Count==0)throw new Exception("Word 未识别原生公式，插入已取消。");
        range.Font.Name="Cambria Math";range.Font.Size=(float)size.Value;
    }
    string CreateEquationDocx(string omml,Dictionary<string,object> data) {
        string file=Path.Combine(root,"temp","equation_"+Guid.NewGuid().ToString("N")+".docx");
        var math=new XmlDocument();math.LoadXml(omml);var mathns=new XmlNamespaceManager(math.NameTable);mathns.AddNamespace("m",M);
        foreach(XmlNode r in math.SelectNodes("//m:r",mathns)){
            var rp=math.CreateElement("w","rPr",W);var fonts=math.CreateElement("w","rFonts",W);fonts.SetAttribute("ascii",W,"Cambria Math");fonts.SetAttribute("hAnsi",W,"Cambria Math");rp.AppendChild(fonts);
            var sz=math.CreateElement("w","sz",W);sz.SetAttribute("val",W,((int)size.Value*2).ToString());rp.AppendChild(sz);
            r.InsertBefore(rp,r.SelectSingleNode("m:t",mathns));
        }
        omml=math.DocumentElement.OuterXml;
        string section="";
        if(data!=null){
            int width=(int)Math.Min(30000,Math.Max(1440,Convert.ToDouble(data["width"])*15+320));
            int height=(int)Math.Min(30000,Math.Max(1440,Convert.ToDouble(data["height"])*15+400));
            section="<w:sectPr><w:pgSz w:w=\""+width+"\" w:h=\""+height+"\"/><w:pgMar w:top=\"120\" w:right=\"120\" w:bottom=\"120\" w:left=\"120\" w:header=\"0\" w:footer=\"0\" w:gutter=\"0\"/></w:sectPr>";
        }
        string xml="<w:document xmlns:w=\""+W+"\" xmlns:m=\""+M+"\"><w:body><w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"+omml+"</w:p>"+section+"</w:body></w:document>";
        using(var zip=ZipFile.Open(file,ZipArchiveMode.Create)){
            WriteZipText(zip,"[Content_Types].xml","<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            WriteZipText(zip,"_rels/.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            WriteZipText(zip,"word/document.xml",xml);
        }
        return file;
    }
    void WriteZipText(ZipArchive zip,string name,string text){using(var writer=new StreamWriter(zip.CreateEntry(name).Open(),new UTF8Encoding(false)))writer.Write(text);}
    void InsertNativeWord(dynamic app,string omml) {
        if(app.Documents.Count==0)throw new Exception("请先打开 Word 文档。");
        dynamic range=app.Selection.Range.Duplicate;range.Collapse(1);
        bool undo=false;try{app.UndoRecord.StartCustomRecord("插入可编辑公式");undo=true;}catch{}bool success=false;
        try{PutWordEquation(range,omml);range.Select();success=true;}
        finally{if(undo)app.UndoRecord.EndCustomRecord();if(!success)try{app.ActiveDocument.Undo(1);}catch{}}
        status.Text="已插入可编辑公式，点击公式内部即可编辑，无需重新生成。";
    }
    XmlDocument PptMath(string omml) {
        var xml=new XmlDocument();xml.LoadXml(omml);
        var ns=new XmlNamespaceManager(xml.NameTable);ns.AddNamespace("m",M);ns.AddNamespace("w",W);
        foreach(XmlNode r in xml.SelectNodes("//m:r",ns)){
            // Office Math in DrawingML uses a:rPr rather than Word's w:rPr.
            var remove=new List<XmlNode>();foreach(XmlNode c in r.ChildNodes)if(c.NamespaceURI==W)remove.Add(c);foreach(var c in remove)r.RemoveChild(c);
            var rp=xml.CreateElement("a","rPr",A);rp.SetAttribute("sz",((int)size.Value*100).ToString());
            var latin=xml.CreateElement("a","latin",A);latin.SetAttribute("typeface","Cambria Math");rp.AppendChild(latin);
            XmlNode text=r.SelectSingleNode("m:t",ns);if(text!=null)r.InsertBefore(rp,text);else r.AppendChild(rp);
        }
        foreach(XmlNode r in xml.SelectNodes("//w:rPr",ns)){var rp=xml.CreateElement("a","rPr",A);rp.SetAttribute("sz",((int)size.Value*100).ToString());r.ParentNode.ReplaceChild(rp,r);}
        return xml;
    }
    string NativePptFile(dynamic app,string omml,Dictionary<string,object> data) {
        string file=Path.Combine(root,"temp","native_"+Guid.NewGuid().ToString("N")+".pptx");
        dynamic scratch=app.Presentations.Add(0);
        try{scratch.Slides.Add(1,12);scratch.SaveAs(file,24);}finally{scratch.Saved=-1;scratch.Close();}
        long cx=(long)Math.Max(914400,Convert.ToDouble(data["width"])*9525+127000);
        long cy=(long)Math.Max(457200,Convert.ToDouble(data["height"])*9525+127000);
        string math=PptMath(omml).DocumentElement.OuterXml;
        string shape="<p:sp xmlns:p=\""+P+"\" xmlns:a=\""+A+"\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\"><p:nvSpPr><p:cNvPr id=\"2\" name=\"Editable equation\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\""+cx+"\" cy=\""+cy+"\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap=\"none\" lIns=\"63500\" tIns=\"63500\" rIns=\"63500\" bIns=\"63500\"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr/><a14:m>"+math+"</a14:m><a:endParaRPr lang=\"en-US\" sz=\""+((int)size.Value*100)+"\"/></a:p></p:txBody></p:sp>";
        using(var zip=ZipFile.Open(file,ZipArchiveMode.Update)){
            var entry=zip.GetEntry("ppt/slides/slide1.xml");string text;using(var reader=new StreamReader(entry.Open()))text=reader.ReadToEnd();
            var slide=new XmlDocument();slide.LoadXml(text);var ns=new XmlNamespaceManager(slide.NameTable);ns.AddNamespace("p",P);
            var tree=slide.SelectSingleNode("//p:spTree",ns);var fragment=slide.CreateDocumentFragment();fragment.InnerXml=shape;tree.AppendChild(fragment);
            entry.Delete();entry=zip.CreateEntry("ppt/slides/slide1.xml");using(var writer=new StreamWriter(entry.Open(),new UTF8Encoding(false)))writer.Write(slide.OuterXml);
        }
        return file;
    }
    void InsertNativePpt(dynamic app,string omml,Dictionary<string,object> data) {
        if(app.Presentations.Count==0)throw new Exception("请先打开 PPT 的目标幻灯片。");
        dynamic dest=app.ActivePresentation;dynamic destWindow=app.ActiveWindow;dynamic slide;try{slide=destWindow.View.Slide;}catch{throw new Exception("请在 PPT 普通视图中打开目标幻灯片。");}
        string file=NativePptFile(app,omml,data);dynamic source=null;string clipboardText=Clipboard.ContainsText()?Clipboard.GetText():null;
        try{
            source=app.Presentations.Open(file,-1,0,-1);source.Windows[1].WindowState=2;
            dynamic pasted=null;Exception last=null;
            for(int attempt=0;attempt<6;attempt++){
                source.Slides[1].Shapes[1].Copy();System.Threading.Thread.Sleep(150);Application.DoEvents();destWindow.Activate();
                try{try{app.StartNewUndoEntry();}catch{}pasted=slide.Shapes.Paste();break;}catch(Exception ex){last=ex;}
            }
            if(pasted==null)throw new Exception("PowerPoint 暂时无法粘贴原生公式，请重试。",last);
            dynamic shape=pasted[1];
            if(shape.HasTextFrame==0)throw new Exception("PowerPoint 未保留原生公式文本框。");
            shape.Left=((float)dest.PageSetup.SlideWidth-(float)shape.Width)/2;shape.Top=((float)dest.PageSetup.SlideHeight-(float)shape.Height)/2;
            shape.Select();status.Text="已插入 PPT 原生公式，点击公式内部可直接编辑，不是图片。";
        }finally{if(source!=null)source.Close();RestoreClipboardText(clipboardText);}
    }
    void SelfTestNative() {
        var report=new List<string>();dynamic doc=null,pres=null;string stamp=DateTime.Now.ToString("yyyyMMdd_HHmmss");
        try{
            wordTest=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));wordTest.Visible=false;doc=wordTest.Documents.Add();
            var samples=new []{@"\tau=\frac{k_p e_q+k_d e_v-z_3}{b_0}",@"q_r(t)=\sum_i A_i v_1(t-t_i)",@"q_d(t)\xrightarrow{\mathrm{TD}}\bigl(v_1(t),v_2(t)\bigr)"};
            foreach(var sample in samples){doc.Range(doc.Content.End-1,doc.Content.End-1).Select();editor.Text=sample;InsertEditable("Word");doc.Range(doc.Content.End-1,doc.Content.End-1).InsertAfter("\r");}
            if(doc.OMaths.Count!=3||doc.InlineShapes.Count!=0)throw new Exception("Word native math count failed");
            doc.OMaths[1].Range.Characters[1].Text="p";if(doc.OMaths.Count!=3)throw new Exception("Word direct editing lost math");
            doc.ExportAsFixedFormat(Path.Combine(root,"output","native_word_"+stamp+".pdf"),17);
            doc.SaveAs2(Path.Combine(root,"temp","native_word_"+stamp+".docx"),16);
            report.Add("PASS Word: 3 native OMaths, no pictures, direct character editing, PDF export");
        }finally{if(doc!=null)doc.Close(0);if(wordTest!=null){if(wordTest.Documents.Count==0)wordTest.Quit(0);wordTest=null;}}
        try{
            pptTest=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));pres=pptTest.Presentations.Add(-1);pptTest.ActiveWindow.WindowState=2;dynamic slide=pres.Slides.Add(1,12);pptTest.ActiveWindow.View.GotoSlide(1);
            editor.Text=@"\tau=\frac{k_p e_q+k_d e_v-z_3}{b_0}";InsertEditable("PowerPoint");
            dynamic shape=slide.Shapes[1];if(shape.HasTextFrame==0)throw new Exception("PPT not a text shape");
            shape.TextFrame2.TextRange.Characters[1,1].Text="p";
            string file=Path.Combine(root,"temp","native_ppt_"+stamp+".pptx");pres.SaveAs(file,24);
            using(var zip=ZipFile.OpenRead(file)){string text;using(var reader=new StreamReader(zip.GetEntry("ppt/slides/slide1.xml").Open()))text=reader.ReadToEnd();if(!text.Contains("oMath")||!text.Contains(":f>"))throw new Exception("PPT native equation structure not persisted");}
            slide.Export(Path.Combine(root,"output","native_ppt_"+stamp+".png"),"PNG",1440,1080);
            report.Add("PASS PPT: editable native math text shape, direct character edit, persisted fraction structure");
        }finally{if(pres!=null){pres.Saved=-1;pres.Close();}if(pptTest!=null){if(pptTest.Presentations.Count==0)pptTest.Quit();pptTest=null;}}
        File.WriteAllLines(Path.Combine(root,"output","native-self-test.txt"),report,Encoding.UTF8);
    }
    void SelfTestNativeScripts(){
        dynamic doc=null;
        try{
            wordTest=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));wordTest.Visible=false;doc=wordTest.Documents.Add();
            editor.Text=@"q_{d,i}^{(2)}+\dot q_r+x^{n+1}";InsertEditable("Word");
            if(doc.OMaths.Count!=1||doc.InlineShapes.Count!=0)throw new Exception("原生公式数量或图片检查失败");
            string xml=(string)doc.OMaths[1].Range.WordOpenXML;
            foreach(string tag in new[]{"sSubSup","sSup","sSub","acc"})if(!xml.Contains(":"+tag+">"))throw new Exception("原生公式缺少结构："+tag);
            doc.ExportAsFixedFormat(Path.Combine(root,"output","native_scripts_validation.pdf"),17);
            File.WriteAllText(Path.Combine(root,"output","native-scripts-test.txt"),"PASS native Word equation retains subscript, superscript, combined scripts and dot accent; no pictures.",Encoding.UTF8);
        }finally{if(doc!=null)doc.Close(0);if(wordTest!=null){if(wordTest.Documents.Count==0)wordTest.Quit(0);wordTest=null;}}
    }
    void SelfTestSpacing(){
        dynamic doc=null;
        try{
            wordTest=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));wordTest.Visible=false;doc=wordTest.Documents.Add();
            editor.Text="\\begin{gathered}\nq_{d,i}^{(2)}=v_1+x^{n+1} \\\\\n\\tau=\\frac{k_p e_q+k_d e_v-z_3}{b_0} \\\\\n\\dot q_r=\\sum_i A_i v_2(t-t_i)\n\\end{gathered}";
            var data=Render();string omml=Omml(data);
            var check=new XmlDocument();check.LoadXml(omml);var ns=new XmlNamespaceManager(check.NameTable);ns.AddNamespace("m",M);
            var rule=check.SelectSingleNode("//m:rSpRule",ns);if(rule==null||rule.Attributes["val",M].Value!="2")throw new Exception("Native row spacing missing");
            InsertNativeWord(wordTest,omml);
            doc.ExportAsFixedFormat(Path.Combine(root,"output","spacing_validation.pdf"),17);
            File.Copy((string)data["png"],Path.Combine(root,"output","spacing_preview.png"),true);
            File.WriteAllText(Path.Combine(root,"output","spacing-test.txt"),"PASS generated rows have 1.2em gaps; legacy batches normalized; native Office equations carry double row-gap rule.",Encoding.UTF8);
        }finally{if(doc!=null)doc.Close(0);if(wordTest!=null){if(wordTest.Documents.Count==0)wordTest.Quit(0);wordTest=null;}}
    }
    void SelfTestWps() {
        string source=@"q_{d,i}^{(2)}=\frac{a_1+b_2}{c^2}";
        var data=LatexTextData(source);
        if((string)data.GetData(DataFormats.UnicodeText)!=source)throw new Exception("LaTeX source changed");
        foreach(string format in new[]{DataFormats.Bitmap,DataFormats.Html,"Embed Source","EnhancedMetafile"})
            if(data.GetDataPresent(format))throw new Exception("Unexpected non-text format: "+format);
        File.WriteAllText(Path.Combine(root,"output","wps-self-test.txt"),"PASS WPS source mode: Unicode LaTeX text only; subscripts, superscripts and fractions preserved; no bitmap, HTML or OLE object formats.",Encoding.UTF8);
    }
}
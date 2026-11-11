using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Collections.Generic;

partial class FormulaHelper : Form {
    [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr h, int id, uint mods, uint key);
    [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr h, int id);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    readonly string root = AppDomain.CurrentDomain.BaseDirectory;
    readonly JavaScriptSerializer json = new JavaScriptSerializer();
    TextBox editor = new TextBox(); PictureBox preview = new PictureBox();
    NumericUpDown size = new NumericUpDown(); Label status = new Label();
    ComboBox target = new ComboBox(); Timer foregroundTimer = new Timer();
    CheckBox nativeMode=new CheckBox();
    string lastTarget="Visio";
    IntPtr lastWpsWindow=IntPtr.Zero;
    Dictionary<string,object> runtime;
    bool busy;
    dynamic wordTest, pptTest;
    public FormulaHelper() {
        Directory.CreateDirectory(Path.Combine(root,"output")); Directory.CreateDirectory(Path.Combine(root,"temp"));
        runtime=json.Deserialize<Dictionary<string,object>>(File.ReadAllText(Path.Combine(root,"runtime.json")));
        Text="通用公式助手 v1.0.0 · GPT 公式提取 / Office 可编辑公式"; Width=1060; Height=700; MinimumSize=new Size(1060,620);
        StartPosition=FormStartPosition.CenterScreen; Font=new Font("Microsoft YaHei UI",10);
        var layout=new TableLayoutPanel {Dock=DockStyle.Fill,RowCount=6,ColumnCount=1,Padding=new Padding(12)};
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48));layout.RowStyles.Add(new RowStyle(SizeType.Percent,52));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute,44));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,88));layout.RowStyles.Add(new RowStyle(SizeType.Percent,48));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48));
        var hint=new Label {Text="要显示成分式、上下标并直接编辑：选好目标软件，点击“复制可编辑公式”，再 Ctrl+V。\n“复制 LaTeX 源码”仅输出代码文本，不会自动变成公式。WPS/Word 请用 .docx，演示请用 .pptx。",Dock=DockStyle.Fill};
        editor.Multiline=true;editor.ScrollBars=ScrollBars.Both;editor.AcceptsReturn=true;editor.WordWrap=false;editor.Dock=DockStyle.Fill;
        editor.Font=new Font("Consolas",13);editor.Text=@"q_r(t)=\sum_i A_i v_1(t-t_i),\qquad \dot q_r(t)=\sum_i A_i v_2(t-t_i)";
        var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(0,8,0,0)};
        target.DropDownStyle=ComboBoxStyle.DropDownList;target.Width=106;target.Items.AddRange(new object[]{"自动识别","Visio","Word","PowerPoint","WPS 文字","WPS 演示"});target.SelectedIndex=0;buttons.Controls.Add(target);
        nativeMode.Checked=true;
        AddButton(buttons,"读取剪贴板",()=>ExtractClipboardLatex(false));
        AddButton(buttons,"预览",()=>{Render();});
        AddButton(buttons,"放大核对",ShowLargeFormula);
        AddButton(buttons,"插入公式",()=>{InsertTarget(false);});
        AddButton(buttons,"复制可编辑公式",()=>{CopyEditableFormula(Target());});
        AddButton(buttons,"复制 LaTeX 源码",CopyLatexSource);
        AddButton(buttons,"读取图形源码",()=>{if(nativeMode.Checked&&Target()!="Visio")throw new Exception("可编辑公式请直接在目标软件中修改。图形源码功能用于 Visio 或未勾选“可编辑”时插入的图形。");LoadTarget();});
        AddButton(buttons,"更新图形公式",()=>{if(nativeMode.Checked&&Target()!="Visio")throw new Exception("可编辑公式无需重新生成，请在 Word/PPT/WPS 中点击公式内部编辑。");InsertTarget(true);});
        buttons.Controls.Add(new Label {Text="字号",AutoSize=true,Margin=new Padding(8,6,0,0)});
        size.Minimum=6;size.Maximum=72;size.Value=14;size.Width=52;buttons.Controls.Add(size);
        preview.Dock=DockStyle.Fill;preview.SizeMode=PictureBoxSizeMode.Zoom;preview.BackColor=Color.White;preview.BorderStyle=BorderStyle.FixedSingle;
        status.Dock=DockStyle.Fill;status.Text="Word：插入光标处；PPT：当前幻灯片；Visio：当前页面。均不自动保存。";
        layout.Controls.Add(hint,0,0);layout.Controls.Add(editor,0,1);layout.Controls.Add(ClipboardControls(),0,2);layout.Controls.Add(buttons,0,3);layout.Controls.Add(preview,0,4);layout.Controls.Add(status,0,5);Controls.Add(layout);
        Shown+=(s,e)=>{if(!RegisterHotKey(Handle,1,0x4003,0x56)) status.Text="快捷键被其他程序占用，请使用窗口内的插入按钮。";StartClipboardExtraction();};
        foregroundTimer.Interval=300;foregroundTimer.Tick+=(s,e)=>{string app=ForegroundTarget();if(app!=null)lastTarget=app;};foregroundTimer.Start();
        FormClosed+=(s,e)=>{foregroundTimer.Stop();foregroundTimer.Dispose();UnregisterHotKey(Handle,1);StopClipboardExtraction();ReleaseNativeClipboardSource();if(preview.Image!=null)preview.Image.Dispose();};
    }
    void AddButton(FlowLayoutPanel panel,string text,Action action) {
        var b=new Button {Text=text,AutoSize=true};b.Click+=(s,e)=>Run(action);panel.Controls.Add(b);
    }
    void Run(Action action) {
        if(busy)return;busy=true;UseWaitCursor=true;
        try{action();}catch(Exception ex){status.Text="操作失败："+ex.Message;File.AppendAllText(Path.Combine(root,"output","errors.log"),DateTime.Now+" "+ex+Environment.NewLine);MessageBox.Show(this,ex.Message,"公式助手",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        finally{busy=false;UseWaitCursor=false;}
    }
    Dictionary<string,object> Render() {
        var output=Path.Combine(root,"output");Directory.CreateDirectory(output);
        var stem=Path.Combine(output,"formula_"+DateTime.Now.ToString("yyyyMMdd_HHmmss_fff"));
        var req=stem+".request.json";
        File.WriteAllText(req,json.Serialize(new {latex=editor.Text,fontPt=(int)size.Value,outputBase=stem}),new UTF8Encoding(false));
        var info=new ProcessStartInfo(Path.GetFullPath(Path.Combine(root,(string)runtime["node"])),"\""+Path.Combine(root,"render.cjs")+"\" \""+req+"\"") {
            UseShellExecute=false,CreateNoWindow=true,RedirectStandardError=true,RedirectStandardOutput=true,WorkingDirectory=root,StandardErrorEncoding=Encoding.UTF8};
        info.EnvironmentVariables["TEMP"]=Path.Combine(root,"temp");info.EnvironmentVariables["TMP"]=Path.Combine(root,"temp");
        using(var p=Process.Start(info)){
            if(!p.WaitForExit(20000)){p.Kill();throw new Exception("公式渲染超过20秒，已停止。请检查公式长度和语法。");}
            var err=p.StandardError.ReadToEnd();if(p.ExitCode!=0)throw new Exception(err);
        }
        var result=json.Deserialize<Dictionary<string,object>>(File.ReadAllText(stem+".json"));
        using(var image=Image.FromFile((string)result["png"])) {var old=preview.Image;preview.Image=new Bitmap(image);if(old!=null)old.Dispose();}
        status.Text="预览完成。插入后为矢量图形；可用“读取选中公式”修改源码。";
        return result;
    }
    dynamic Visio() {
        try{return Marshal.GetActiveObject("Visio.Application");}catch{throw new Exception("请先打开 Visio 和需要插入公式的图纸。");}
    }
    string ForegroundTarget() {
        uint pid;GetWindowThreadProcessId(GetForegroundWindow(),out pid);
        try{string name=Process.GetProcessById((int)pid).ProcessName.ToUpperInvariant();if(name=="VISIO")return "Visio";if(name=="WINWORD")return "Word";if(name=="POWERPNT")return "PowerPoint";if(name=="WPS"||name=="KWPS"){lastWpsWindow=GetForegroundWindow();return "WPS 文字";}if(name=="WPP"||name=="KWPP"){lastWpsWindow=GetForegroundWindow();return "WPS 演示";}}catch{}
        return null;
    }
    string Target() {return target.SelectedIndex>0?target.SelectedItem.ToString():lastTarget;}
    dynamic Office(string name) {
        if(name=="Word"&&wordTest!=null)return wordTest;if(name=="PowerPoint"&&pptTest!=null)return pptTest;
        string progid=name=="WPS 文字"?"KWPS.Application":name=="WPS 演示"?"KWPP.Application":name+".Application";
        try{return Marshal.GetActiveObject(progid);}catch{
            if(name.StartsWith("WPS",StringComparison.Ordinal))try{return Activator.CreateInstance(Type.GetTypeFromProgID(progid));}catch{}
            throw new Exception("请先打开 "+name+" 和要编辑的文件，然后在助手中选择对应软件。");}
    }
    string Pack(Dictionary<string,object> data) {return "OfficeFormulaHelper:"+json.Serialize(new {latex=(string)data["latex"],fontPt=(int)size.Value});}
    Dictionary<string,object> Unpack(string value) {
        const string prefix="OfficeFormulaHelper:";
        if(value==null||!value.StartsWith(prefix,StringComparison.Ordinal))throw new Exception("请选择由本助手插入的公式。");
        return json.Deserialize<Dictionary<string,object>>(value.Substring(prefix.Length));
    }
    void SetEditor(Dictionary<string,object> data) {
        editor.Text=(string)data["latex"];size.Value=Math.Max(6,Math.Min(72,Convert.ToDecimal(data["fontPt"])));
        status.Text="已读取 "+Target()+" 选中公式，修改后点击“更新选中公式”。";
    }
    void InsertTarget(bool replace) {string app=Target();if(nativeMode.Checked&&!replace){InsertEditable(app);return;}if(app=="Word")InsertWord(replace);else if(app=="PowerPoint")InsertPpt(replace);else Insert(replace);}
    void LoadTarget() {
        string name=Target();if(name=="Visio"){LoadSelected();return;}
        dynamic app=Office(name);
        if(name=="Word"){
            if(app.Documents.Count==0||app.Selection.InlineShapes.Count!=1)throw new Exception("请在 Word 中单击选中一个本助手插入的嵌入式公式图片。");
            SetEditor(Unpack((string)app.Selection.InlineShapes[1].AlternativeText));
        }else{
            if(app.Presentations.Count==0||app.ActiveWindow.Selection.Type!=2||app.ActiveWindow.Selection.ShapeRange.Count!=1)throw new Exception("请在 PPT 中选中一个本助手插入的公式。");
            SetEditor(Unpack((string)app.ActiveWindow.Selection.ShapeRange[1].AlternativeText));
        }
    }
    void InsertWord(bool replace) {
        dynamic app=Office("Word");if(app.Documents.Count==0)throw new Exception("请先打开 Word 文档。");
        dynamic doc=app.ActiveDocument;dynamic old=null;dynamic range=app.Selection.Range.Duplicate;
        if(replace){if(app.Selection.InlineShapes.Count!=1)throw new Exception("请在 Word 中选中要更新的公式图片。");old=app.Selection.InlineShapes[1];Unpack((string)old.AlternativeText);range=old.Range.Duplicate;}
        range.Collapse(1); // insert at selection start; never overwrite selected prose
        var data=Render();dynamic shape=null;bool committed=false;bool changed=false;
        app.UndoRecord.StartCustomRecord(replace?"更新公式":"插入公式");
        try{
            shape=doc.InlineShapes.AddPicture((string)data["svg"],false,true,range);changed=true;
            double width=Convert.ToDouble(data["width"])*0.75,height=Convert.ToDouble(data["height"])*0.75;
            double available=(double)range.Sections[1].PageSetup.PageWidth-(double)range.Sections[1].PageSetup.LeftMargin-(double)range.Sections[1].PageSetup.RightMargin;
            if(width>available&&available>0){height*=available/width;width=available;}
            shape.LockAspectRatio=0;shape.Width=(float)width;shape.Height=(float)height;shape.LockAspectRatio=-1;
            shape.AlternativeText=Pack(data);shape.Title="LaTeX 公式";
            if(replace)old.Delete();shape.Select();committed=true;
        }finally{app.UndoRecord.EndCustomRecord();if(!committed&&changed)try{doc.Undo(1);}catch{}}
        status.Text="已插入 Word 光标处，公式随正文排版。Ctrl+Z 可撤销；文档未自动保存。";
    }
    void InsertPpt(bool replace) {
        dynamic app=Office("PowerPoint");if(app.Presentations.Count==0)throw new Exception("请先打开 PPT 并切换到要插入的幻灯片。");
        dynamic window=app.ActiveWindow;dynamic slide;
        try{slide=window.View.Slide;}catch{throw new Exception("请在 PPT 普通视图中选中一个幻灯片页面。");}
        dynamic old=null;float left=0,top=0,rotation=0;
        if(replace){if(window.Selection.Type!=2||window.Selection.ShapeRange.Count!=1)throw new Exception("请在 PPT 中选中一个要更新的公式。");old=window.Selection.ShapeRange[1];Unpack((string)old.AlternativeText);left=(float)old.Left;top=(float)old.Top;rotation=(float)old.Rotation;}
        var data=Render();float width=(float)(Convert.ToDouble(data["width"])*0.75),height=(float)(Convert.ToDouble(data["height"])*0.75);
        float sw=(float)app.ActivePresentation.PageSetup.SlideWidth,sh=(float)app.ActivePresentation.PageSetup.SlideHeight;
        float scale=Math.Min(1f,Math.Min((sw-24)/width,(sh-24)/height));width*=scale;height*=scale;
        if(!replace){left=(sw-width)/2;top=(sh-height)/2;}
        app.StartNewUndoEntry();dynamic shape=null;bool committed=false;
        try{
            shape=slide.Shapes.AddPicture((string)data["svg"],0,-1,left,top,width,height);
            shape.AlternativeText=Pack(data);shape.Title="LaTeX 公式";shape.Rotation=rotation;
            if(replace)old.Delete();shape.Select();committed=true;
        }finally{if(!committed&&shape!=null)try{shape.Delete();}catch{}}
        status.Text="已插入 PPT 当前幻灯片，可拖动、缩放。Ctrl+Z 可撤销；演示文稿未自动保存。";
    }
    void LoadSelected() {
        dynamic app=Visio();dynamic sel=app.ActiveWindow.Selection;
        if(sel.Count!=1)throw new Exception("请在 Visio 中选中一个由本助手插入的公式。");
        dynamic shape=sel[1];
        if(shape.CellExistsU["User.EquationLaTeX",0]==0)throw new Exception("选中的图形没有本助手保存的公式源码。");
        editor.Text=(string)shape.CellsU["User.EquationLaTeX"].ResultStrU[0];
        if(shape.CellExistsU["User.EquationFontPt",0]!=0)size.Value=Math.Max(6,Math.Min(72,(decimal)(double)shape.CellsU["User.EquationFontPt"].ResultIU));
        status.Text="已读取选中公式，修改后点击“更新选中公式”。";
    }
    void Insert(bool replace) {
        dynamic app=Visio();if(app.ActiveDocument==null||app.ActivePage==null)throw new Exception("请先在 Visio 中打开一个绘图页面。");
        dynamic old=null;double x=0,y=0;
        if(replace){dynamic selection=app.ActiveWindow.Selection;if(selection.Count!=1)throw new Exception("请先选中要更新的一个公式。");old=selection[1];if(old.CellExistsU["User.EquationLaTeX",0]==0)throw new Exception("只更新由本助手插入的公式，避免替换其他图形。");x=(double)old.CellsU["PinX"].ResultIU;y=(double)old.CellsU["PinY"].ResultIU;}
        var data=Render();dynamic page=app.ActivePage;
        int scope=app.BeginUndoScope(replace?"更新 LaTeX 公式":"插入 LaTeX 公式");bool success=false;
        try{
            dynamic shape=page.Import((string)data["svg"]);
            shape.CellsU["Width"].ResultIU=Convert.ToDouble(data["width"])/96;
            shape.CellsU["Height"].ResultIU=Convert.ToDouble(data["height"])/96;
            if(!replace){x=(double)page.PageSheet.CellsU["PageWidth"].ResultIU/2;y=(double)page.PageSheet.CellsU["PageHeight"].ResultIU/2;}
            shape.CellsU["PinX"].ResultIU=x;shape.CellsU["PinY"].ResultIU=y;
            if(shape.SectionExists[242,0]==0)shape.AddSection(242);
            shape.AddNamedRow(242,"EquationLaTeX",0);shape.CellsU["User.EquationLaTeX"].FormulaU="\""+((string)data["latex"]).Replace("\"","\"\"")+"\"";
            shape.AddNamedRow(242,"EquationFontPt",0);shape.CellsU["User.EquationFontPt"].FormulaU=((int)size.Value).ToString(System.Globalization.CultureInfo.InvariantCulture);
            if(replace){shape.CellsU["Angle"].FormulaU=old.CellsU["Angle"].FormulaU;old.Delete();}
            app.ActiveWindow.DeselectAll();app.ActiveWindow.Select(shape,2);success=true;
            status.Text="已插入当前 Visio 页面，可拖动和缩放；Ctrl+Z 可撤销。图纸尚未自动保存。";
        }finally{app.EndUndoScope(scope,success);}
    }
    protected override void WndProc(ref Message m) {
        if(m.Msg==0x031D&&IsHandleCreated&&!IsDisposed)BeginInvoke((Action)AutomaticClipboardExtraction);
        if(m.Msg==0x0312&&m.WParam.ToInt32()==2)Run(()=>ExtractClipboardLatex(true));
        if(m.Msg==0x0312&&m.WParam.ToInt32()==1){
            string app=ForegroundTarget();
            if(app!=null)Run(()=>{lastTarget=app;ExtractClipboardLatex(false);if(nativeMode.Checked)InsertEditable(app);else if(app=="Word")InsertWord(false);else if(app=="PowerPoint")InsertPpt(false);else Insert(false);});
        }
        base.WndProc(ref m);
    }
    void SelfTest() {
        dynamic app=Visio();dynamic previous=app.ActiveWindow;dynamic doc=null;
        var report=new List<string>();
        try {
            doc=app.Documents.Add("");dynamic page=doc.Pages[1];page.Name="TD-ZVD-LADRC 公式示例";
            var formulas=new [] {
                @"q_d(t)=q_0+(q_f-q_0)(10\tau^3-15\tau^4+6\tau^5),\quad \tau=t/T",
                @"q_d(t)\xrightarrow{\mathrm{TD}}\bigl(v_1(t),v_2(t)\bigr)",
                @"q_r(t)=\sum_i A_i v_1(t-t_i)",
                @"\dot q_r(t)=\sum_i A_i v_2(t-t_i)",
                @"e_q=q_r-\hat q,\qquad e_v=\dot q_r-\hat{\dot q}",
                @"\tau=\frac{k_p e_q+k_d e_v-z_3}{b_0}"
            };
            for(int i=0;i<formulas.Length;i++) {
                editor.Text=formulas[i];Insert(false);dynamic shape=app.ActiveWindow.Selection[1];
                shape.CellsU["PinX"].ResultIU=4.0;shape.CellsU["PinY"].ResultIU=10.0-i*1.3;
                string read=(string)shape.CellsU["User.EquationLaTeX"].ResultStrU[0];
                if(read!=formulas[i])throw new Exception("源码回读不一致 "+i);
                report.Add("PASS SVG import, source roundtrip: "+i);
            }
            LoadSelected();editor.Text=@"\tau=\frac{u_0-\hat f}{b_0}";Insert(true);
            if(page.Shapes.Count!=formulas.Length)throw new Exception("更新后图形数量异常");
            report.Add("PASS update selected equation and preserve shape count");
            app.Undo();if(page.Shapes.Count!=formulas.Length)throw new Exception("撤销后图形数量异常");
            report.Add("PASS single-step undo");
            string output=Path.Combine(root,"output","公式示例_"+DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            doc.SaveAs(output+".vsdx");page.Export(output+".png");
            report.Add("Saved example: "+output+".vsdx");
            File.WriteAllLines(Path.Combine(root,"output","self-test.txt"),report,Encoding.UTF8);
        }finally{if(doc!=null){doc.Saved=true;doc.Close();}if(previous!=null)previous.Activate();}
    }
    void SelfTestOffice() {
        var report=new List<string>();dynamic wdoc=null,pres=null;
        string stamp=DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var samples=new [] {@"q_d(t)\xrightarrow{\mathrm{TD}}\bigl(v_1(t),v_2(t)\bigr)",@"q_r(t)=\sum_i A_i v_1(t-t_i)",@"\dot q_r(t)=\sum_i A_i v_2(t-t_i)",@"\tau=\frac{k_p e_q+k_d e_v-z_3}{b_0}"};
        try{
            wordTest=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            wordTest.Visible=false;wdoc=wordTest.Documents.Add();wdoc.Content.Text="KEEP_SELECTED_TEXT\r";
            target.SelectedItem="Word";
            wdoc.Range(0,18).Select();editor.Text=samples[0];InsertWord(false);
            if(!((string)wdoc.Content.Text).Contains("KEEP_SELECTED_TEXT"))throw new Exception("Word overwritten selected prose");
            LoadTarget();if(editor.Text!=samples[0])throw new Exception("Word source mismatch");
            report.Add("PASS Word: SVG insertion, source readback, selected prose preserved");
            editor.Text=samples[3];InsertWord(true);if(wdoc.InlineShapes.Count!=1)throw new Exception("Word update count mismatch");
            LoadTarget();if(editor.Text!=samples[3])throw new Exception("Word update source mismatch");
            wdoc.Undo(1);if(wdoc.InlineShapes.Count!=1)throw new Exception("Word undo count mismatch");
            wdoc.InlineShapes[1].Select();LoadTarget();if(editor.Text!=samples[0])throw new Exception("Word undo source mismatch");
            report.Add("PASS Word: update, source persistence, one-step undo");
            for(int i=1;i<samples.Length;i++){dynamic end=wdoc.Range(wdoc.Content.End-1,wdoc.Content.End-1);end.InsertAfter("\r");wdoc.Range(wdoc.Content.End-1,wdoc.Content.End-1).Select();editor.Text=samples[i];InsertWord(false);}
            wdoc.ExportAsFixedFormat(Path.Combine(root,"output","word_validation_"+stamp+".pdf"),17);
            report.Add("PASS Word: four formula types rendered to PDF");
            string file=Path.Combine(root,"temp","word_validation_"+stamp+".docx");wdoc.SaveAs2(file,16);wdoc.Close(0);wdoc=null;wdoc=wordTest.Documents.Open(file);wdoc.InlineShapes[1].Select();LoadTarget();if(editor.Text!=samples[0])throw new Exception("Word saved source mismatch");
            report.Add("PASS Word: source survives save and reopen");
        }finally{if(wdoc!=null)wdoc.Close(0);if(wordTest!=null){if(wordTest.Documents.Count==0)wordTest.Quit(0);wordTest=null;}}
        try{
            pptTest=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
            pres=pptTest.Presentations.Add(-1);pptTest.ActiveWindow.WindowState=2;
            dynamic slide=pres.Slides.Add(1,12);pptTest.ActiveWindow.View.GotoSlide(1);target.SelectedItem="PowerPoint";
            for(int i=0;i<samples.Length;i++){
                editor.Text=samples[i];InsertPpt(false);LoadTarget();if(editor.Text!=samples[i])throw new Exception("PPT source mismatch");
                dynamic shape=pptTest.ActiveWindow.Selection.ShapeRange[1];shape.Left=70f;shape.Top=40f+i*85f;
            }
            report.Add("PASS PPT: SVG insertion and source readback for four formulas");
            editor.Text=@"e_q=q_r-\hat q";InsertPpt(true);if(slide.Shapes.Count!=samples.Length)throw new Exception("PPT update count mismatch");
            LoadTarget();if(editor.Text!=@"e_q=q_r-\hat q")throw new Exception("PPT updated source mismatch");
            report.Add("PASS PPT: update selected formula and preserve shape count");
            pptTest.StartNewUndoEntry();pptTest.CommandBars.ExecuteMso("Undo");
            if(slide.Shapes.Count!=samples.Length)throw new Exception("PPT undo count mismatch");slide.Shapes[slide.Shapes.Count].Select();LoadTarget();if(editor.Text!=samples[3])throw new Exception("PPT undo source mismatch");
            report.Add("PASS PPT: one-step undo restores previous formula");
            slide.Export(Path.Combine(root,"output","ppt_validation_"+stamp+".png"),"PNG",1440,1080);
            // Check persisted metadata after save/reopen without touching any user presentation.
            string file=Path.Combine(root,"temp","ppt_validation_"+stamp+".pptx");pres.SaveAs(file,24);
            pres.Close();pres=null;pres=pptTest.Presentations.Open(file,0,0,-1);pptTest.ActiveWindow.View.GotoSlide(1);
            pres.Slides[1].Shapes[1].Select();LoadTarget();if(editor.Text!=samples[0])throw new Exception("PPT saved source mismatch");
            report.Add("PASS PPT: source survives save and reopen");
        }finally{if(pres!=null){pres.Saved=-1;pres.Close();}if(pptTest!=null){if(pptTest.Presentations.Count==0)pptTest.Quit();pptTest=null;}}
        File.WriteAllLines(Path.Combine(root,"output","office-self-test.txt"),report,Encoding.UTF8);
    }
    [STAThread] static void Main(string[] args) {
        Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
        try{if(args.Length>0&&args[0]=="--self-test"){using(var f=new FormulaHelper())f.SelfTest();}else if(args.Length>0&&args[0]=="--self-test-office"){using(var f=new FormulaHelper())f.SelfTestOffice();}else if(args.Length>0&&args[0]=="--self-test-native"){using(var f=new FormulaHelper())f.SelfTestNative();}else if(args.Length>0&&args[0]=="--self-test-wps"){using(var f=new FormulaHelper())f.SelfTestWps();}else if(args.Length>0&&args[0]=="--self-test-extraction"){using(var f=new FormulaHelper())f.SelfTestExtraction();}else if(args.Length>0&&args[0]=="--self-test-spacing"){using(var f=new FormulaHelper())f.SelfTestSpacing();}else if(args.Length>0&&args[0]=="--self-test-scripts"){using(var f=new FormulaHelper())f.SelfTestNativeScripts();}else if(args.Length>0&&args[0]=="--test-copied-formula"){using(var f=new FormulaHelper())f.TestUserCopiedFormula();}else Application.Run(new FormulaHelper());}
        catch(Exception ex){if(args.Length>0){File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"output","self-test-error.txt"),ex.ToString());Environment.ExitCode=1;}else MessageBox.Show(ex.Message,"Visio 公式助手");}
    }
}

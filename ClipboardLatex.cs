using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

partial class FormulaHelper {
    [DllImport("user32.dll")] static extern bool AddClipboardFormatListener(IntPtr window);
    [DllImport("user32.dll")] static extern bool RemoveClipboardFormatListener(IntPtr window);
    [DllImport("user32.dll")] static extern uint GetClipboardSequenceNumber();
    CheckBox autoLatex=new CheckBox();
    bool handlingClipboard;
    uint writtenClipboardSequence;
    class ExtractedFormula {
        public string Latex,Origin;
        public bool Structured;
        public override string ToString(){string s=Latex.Replace("\r"," ").Replace("\n"," ");return s.Length>105?s.Substring(0,105)+"…":s;}
    }
    FlowLayoutPanel ClipboardControls(){
        var bar=new FlowLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(0,5,0,0)};
        AddButton(bar,"提取并复制 LaTeX",()=>ExtractClipboardLatex(true));
        autoLatex.Text="复制内容自动提取公式（过滤说明文字）";autoLatex.AutoSize=true;autoLatex.Checked=true;autoLatex.Margin=new Padding(12,6,6,0);bar.Controls.Add(autoLatex);
        bar.Controls.Add(new Label {Text="Ctrl+Alt+L 可手动选取；多条公式自动合并为多行",AutoSize=true,Margin=new Padding(8,7,0,0)});
        return bar;
    }
    void StartClipboardExtraction(){
        if(!AddClipboardFormatListener(Handle))status.Text="自动提取未启用，可使用“提取并复制 LaTeX”按钮。";
        if(!RegisterHotKey(Handle,2,0x4003,0x4c))status.Text="Ctrl+Alt+L 已被占用，请使用提取按钮。";
    }
    void StopClipboardExtraction(){RemoveClipboardFormatListener(Handle);UnregisterHotKey(Handle,2);}
    static string DecodeHtml(string s){return WebUtility.HtmlDecode(s);}
    static string StripHtml(string s){return DecodeHtml(Regex.Replace(s,"<[^>]+>",""));}
    static string CleanLatex(string s){
        s=s.Trim().Trim('\0');
        var code=Regex.Match(s,@"\A```(?:latex|tex|math|text)?\s*\r?\n(?<v>[\s\S]*?)\r?\n```\z",RegexOptions.IgnoreCase);
        if(code.Success)s=code.Groups["v"].Value.Trim();
        foreach(var pair in new[]{new[]{"\\[","\\]"},new[]{"\\(","\\)"},new[]{"$$","$$"},new[]{"$","$"}})
            if(s.Length>=pair[0].Length+pair[1].Length&&s.StartsWith(pair[0],StringComparison.Ordinal)&&s.EndsWith(pair[1],StringComparison.Ordinal)){s=s.Substring(pair[0].Length,s.Length-pair[0].Length-pair[1].Length).Trim();break;}
        return s;
    }
    static bool WrappedFormula(string s){s=s.Trim();return (s.StartsWith("\\[")&&s.EndsWith("\\]"))||(s.StartsWith("\\(")&&s.EndsWith("\\)"))||(s.StartsWith("$")&&s.EndsWith("$")&&s.Length>2);}
    static List<ExtractedFormula> FindFormulas(string html,string plain){
        html=html??"";plain=plain??"";
        if(html.Length>2000000||plain.Length>1000000)throw new Exception("复制内容过长，请只选择公式或较短的回答片段。");
        var result=new List<ExtractedFormula>();var seen=new HashSet<string>(StringComparer.Ordinal);
        Action<string,string,bool> add=(value,origin,structured)=>{string s=CleanLatex(value);if(s.Length>0&&s.Length<=16000&&seen.Add(s))result.Add(new ExtractedFormula {Latex=s,Origin=origin,Structured=structured});};
        // Codex currently emits data-math-source with the original LaTeX in CF_HTML.
        foreach(Match match in Regex.Matches(html,@"\b(?:data-math-source|data-latex|data-tex|alttext)\s*=\s*(?<q>[""'])(?<v>[\s\S]*?)\k<q>",RegexOptions.IgnoreCase))
            add(DecodeHtml(match.Groups["v"].Value),"HTML 中的公式源码",true);
        foreach(Match match in Regex.Matches(html,@"<(?:[\w-]+:)?annotation\b(?<a>[^>]*)>(?<v>[\s\S]*?)</(?:[\w-]+:)?annotation\s*>",RegexOptions.IgnoreCase))
            if(Regex.IsMatch(match.Groups["a"].Value,@"\bencoding\s*=\s*[""'](?:application/(?:x-)?(?:tex|latex)|TeX)[""']",RegexOptions.IgnoreCase))add(StripHtml(match.Groups["v"].Value),"MathML 的 TeX 注释",true);
        foreach(Match match in Regex.Matches(html,@"<script\b(?<a>[^>]*)>(?<v>[\s\S]*?)</script\s*>",RegexOptions.IgnoreCase))
            if(Regex.IsMatch(match.Groups["a"].Value,@"\btype\s*=\s*[""']math/tex(?:;[^""']*)?[""']",RegexOptions.IgnoreCase))add(DecodeHtml(match.Groups["v"].Value),"MathJax TeX 源码",true);
        if(result.Count>0)return result;
        // Search explicit math delimiters; do not infer fractions from flattened visible text.
        foreach(Match match in Regex.Matches(plain,@"\\\[(?<v>[\s\S]*?)\\\]|\\\((?<v>[\s\S]*?)\\\)|\$\$(?<v>[\s\S]*?)\$\$|(?<![\\$])\$(?!\$)(?<v>[^\r\n$]+)\$(?!\$)"))add(match.Groups["v"].Value,"文本中的公式标记",false);
        if(result.Count>0)return result;
        string cleaned=CleanLatex(plain);
        bool looksLatex=Regex.IsMatch(cleaned,@"\\(?:frac|sum|prod|int|sqrt|dot|hat|begin|mathrm|mathbf|left|right|tau|omega|alpha|beta|theta|xrightarrow)\b")||Regex.IsMatch(cleaned,@"[_^]\{");
        bool explicitScripts=Regex.IsMatch(cleaned,@"[_^](?:\{|[A-Za-z0-9])");
        if(looksLatex||explicitScripts)add(cleaned,"纯文本 LaTeX（请核对上下标）",false);
        return result;
    }
    static string ClipboardString(IDataObject data,string format){
        if(data==null||!data.GetDataPresent(format,false))return "";
        object value=data.GetData(format,false);if(value is string)return (string)value;
        var stream=value as MemoryStream;if(stream!=null&&stream.Length<=2000000)return Encoding.UTF8.GetString(stream.ToArray()).TrimEnd('\0');
        return "";
    }
    List<ExtractedFormula> ClipboardFormulas(out string plain){
        var data=Clipboard.GetDataObject();plain=ClipboardString(data,DataFormats.UnicodeText);
        string html=ClipboardString(data,DataFormats.Html);
        foreach(string format in new[]{"application/x-tex","application/x-latex","TeX","LaTeX"}){
            string value=ClipboardString(data,format);if(!string.IsNullOrWhiteSpace(value))return new List<ExtractedFormula>{new ExtractedFormula {Latex=CleanLatex(value),Origin="剪贴板 TeX 格式",Structured=true}};
        }
        return FindFormulas(html,plain);
    }
    ExtractedFormula ChooseFormula(List<ExtractedFormula> values){
        using(var dialog=new Form {Text="选择要复制的公式",Width=790,Height=460,StartPosition=FormStartPosition.CenterParent,MinimizeBox=false,MaximizeBox=false}){
            var panel=new TableLayoutPanel {Dock=DockStyle.Fill,RowCount=3,ColumnCount=1,Padding=new Padding(10)};
            panel.RowStyles.Add(new RowStyle(SizeType.Percent,55));panel.RowStyles.Add(new RowStyle(SizeType.Percent,45));panel.RowStyles.Add(new RowStyle(SizeType.Absolute,42));
            var list=new ListBox {Dock=DockStyle.Fill};foreach(var value in values)list.Items.Add(value);
            var text=new TextBox {Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,Dock=DockStyle.Fill,Font=editor.Font};
            list.SelectedIndexChanged+=(s,e)=>{text.Text=((ExtractedFormula)list.SelectedItem).Latex;};list.SelectedIndex=0;
            var ok=new Button {Text="复制所选 LaTeX",Dock=DockStyle.Right,Width=145,DialogResult=DialogResult.OK};
            list.DoubleClick+=(s,e)=>{dialog.DialogResult=DialogResult.OK;dialog.Close();};
            panel.Controls.Add(list,0,0);panel.Controls.Add(text,0,1);panel.Controls.Add(ok,0,2);dialog.Controls.Add(panel);dialog.AcceptButton=ok;
            return dialog.ShowDialog(this)==DialogResult.OK?(ExtractedFormula)list.SelectedItem:null;
        }
    }
    void SetLatexClipboard(ExtractedFormula value){
        editor.Text=value.Latex;Clipboard.SetDataObject(LatexTextData(value.Latex),true,10,80);writtenClipboardSequence=GetClipboardSequenceNumber();
        status.Text="已复制 LaTeX（"+value.Origin+"）。"+ScriptSummary(value.Latex);
    }
    static DataObject LatexTextData(string source){var data=new DataObject();data.SetText(source,TextDataFormat.UnicodeText);return data;}
    void CopyLatexSource(){
        string source=CleanLatex(editor.Text);if(string.IsNullOrWhiteSpace(source))throw new Exception("请先提取公式或在编辑框中输入 LaTeX。");
        SetLatexClipboard(new ExtractedFormula {Latex=source,Origin="编辑框中的源码文本",Structured=true});
        status.Text="已复制纯 LaTeX 文本。到 WPS / Word / PPT 按 Ctrl+V 可直接修改代码，不会粘贴成图片。";
    }
    void ExtractClipboardLatex(bool copy){
        string plain;var found=ClipboardFormulas(out plain);
        if(found.Count==0)throw new Exception("没有找到公式源码。请在 GPT/Codex 中完整选中公式后按 Ctrl+C，再提取。只有截图或丢失结构的普通字符时，无法可靠还原 LaTeX。");
        var chosen=found.Count==1?found[0]:ChooseFormula(found);if(chosen==null)return;
        if(copy)SetLatexClipboard(chosen);else{editor.Text=chosen.Latex;status.Text="已读取公式（"+chosen.Origin+"）。";}
    }
    void AutomaticClipboardExtraction(){
        if(!autoLatex.Checked||busy||handlingClipboard||GetClipboardSequenceNumber()==writtenClipboardSequence)return;
        handlingClipboard=true;
        try{
            string plain;var found=ClipboardFormulas(out plain);
            // User explicitly requested stripping prose surrounding copied formulas.
            var value=AutomaticValue(found);if(value!=null)SetLatexClipboard(value);
        }catch(ExternalException){}catch(Exception ex){status.Text="未自动提取："+ex.Message;}
        finally{handlingClipboard=false;}
    }
    static ExtractedFormula AutomaticValue(List<ExtractedFormula> values){
        var reliable=new List<ExtractedFormula>();foreach(var v in values)if(v.Structured||v.Origin=="文本中的公式标记")reliable.Add(v);
        if(reliable.Count==0)return null;if(reliable.Count==1)return reliable[0];
        var lines=new List<string>();foreach(var value in reliable)lines.Add(value.Latex);
        string latex="\\begin{gathered}\n"+string.Join(" \\\\[1.2em]\n",lines)+"\n\\end{gathered}";
        if(latex.Length>16000)throw new Exception("提取出的公式较多，请缩小复制范围或使用手动选择。");
        return new ExtractedFormula {Latex=latex,Origin="按原顺序提取 "+reliable.Count+" 条公式并过滤文字",Structured=true};
    }
    string ScriptSummary(string latex){return "下标标记 "+Regex.Matches(latex,@"(?<!\\)_").Count+" 处，上标标记 "+Regex.Matches(latex,@"(?<!\\)\^").Count+" 处。可点“放大核对”检查。";}
    void ShowLargeFormula(){
        Render();
        using(var dialog=new Form {Text="公式放大核对 · 上下标以源码为准",Width=1100,Height=570,StartPosition=FormStartPosition.CenterParent}){
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,RowCount=3,ColumnCount=1,Padding=new Padding(12)};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,40));layout.RowStyles.Add(new RowStyle(SizeType.Percent,65));layout.RowStyles.Add(new RowStyle(SizeType.Percent,35));
            var label=new Label {Text=ScriptSummary(editor.Text),Dock=DockStyle.Fill};
            var picture=new PictureBox {Dock=DockStyle.Fill,BackColor=System.Drawing.Color.White,SizeMode=PictureBoxSizeMode.Zoom,Image=new System.Drawing.Bitmap(preview.Image)};
            var source=new TextBox {Text=editor.Text,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,Font=new System.Drawing.Font("Consolas",16),Dock=DockStyle.Fill};
            layout.Controls.Add(label,0,0);layout.Controls.Add(picture,0,1);layout.Controls.Add(source,0,2);dialog.Controls.Add(layout);
            dialog.ShowDialog(this);picture.Image.Dispose();
        }
    }
    void SelfTestExtraction(){
        var report=new List<string>();Action<bool,string> check=(ok,name)=>{if(!ok)throw new Exception("提取测试失败："+name);report.Add("PASS "+name);};
        string sample=@"q_d(t)\xrightarrow{\mathrm{TD}}\bigl(v_1(t),v_2(t)\bigr)";
        var first=FindFormulas("<span data-math-source=\""+sample+"\" data-math-display=\"true\">visible</span>","\\[\n"+sample+"\n\\]");check(first.Count==1&&first[0].Latex==sample&&first[0].Structured,"Codex data-math-source exact extraction");
        var escaped=FindFormulas("<span data-math-source='a&amp;b&#92;quad c' />","");check(escaped[0].Latex==@"a&b\quad c","HTML entities decoded once");
        var annotation=FindFormulas("<math><semantics><mrow/><annotation encoding='application/x-tex'>\\frac{a}{b}</annotation></semantics></math>","");check(annotation[0].Latex==@"\frac{a}{b}","KaTeX MathML annotation");
        var many=FindFormulas("<i data-math-source='x^2'></i><i data-math-source='y^2'></i><i data-math-source='x^2'></i>","");check(many.Count==2,"Multiple formulas preserve order and deduplicate");
        check(FindFormulas("",@"说明 \[q_r=\sum_i A_i q_i\] 然后 \(e=q_r-q\)").Count==2,"Delimited formulas in copied response");
        check(FindFormulas("","这是一段普通说明文字").Count==0,"Ordinary prose rejected");
        check(FindFormulas("","").Count==0,"Image-only clipboard rejected");
        string scripts=@"q_{d,i}^{(2)}+v_1+x^{n+1}+\dot q_r";
        var scripted=FindFormulas("<span data-math-source='"+scripts+"'></span>",scripts);check(scripted[0].Latex==scripts,"Nested and multi-character scripts preserved exactly");
        check(FindFormulas("","qr(t)=v1(t)+x2").Count==0,"Flattened ambiguous subscripts are not guessed");
        check(FindFormulas("",@"v_1=x^2")[0].Latex==@"v_1=x^2","Single-character plain LaTeX scripts preserved");
        var mixed=FindFormulas("前面的说明<span data-math-source='q_{d,i}^{(2)}'></span>后面的说明","前面的说明和后面的说明");
        check(AutomaticValue(mixed).Latex==@"q_{d,i}^{(2)}","Mixed prose stripped without changing scripts");
        var merged=AutomaticValue(many);check(merged.Latex.Contains(@"\begin{gathered}")&&merged.Latex.IndexOf("x^2",StringComparison.Ordinal)<merged.Latex.IndexOf("y^2",StringComparison.Ordinal),"Multiple formulas automatically merged in order");
        check(merged.Latex.Contains(@"\\[1.2em]"),"Generated formula rows have explicit vertical gaps");
        File.WriteAllText(Path.Combine(root,"temp","merged-extraction.request.json"),json.Serialize(new {latex=merged.Latex,fontPt=18,outputBase=Path.Combine(root,"output","merged-extraction")}),new UTF8Encoding(false));
        check(AutomaticValue(FindFormulas("","普通文字"))==null,"Text with no math leaves clipboard untouched");
        var textOnly=LatexTextData(scripts);check((string)textOnly.GetData(DataFormats.UnicodeText)==scripts&&!textOnly.GetDataPresent(DataFormats.Bitmap)&&!textOnly.GetDataPresent(DataFormats.Html)&&!textOnly.GetDataPresent("Embed Source"),"LaTeX copy provides only editable text, no images or embedded objects");
        File.WriteAllLines(Path.Combine(root,"output","extraction-self-test.txt"),report,Encoding.UTF8);
    }
    void TestUserCopiedFormula(){
        string plain;var values=ClipboardFormulas(out plain);string expected=@"q_d(t)\xrightarrow{\mathrm{TD}}\bigl(v_1(t),v_2(t)\bigr)";
        if(values.Count!=1||values[0].Latex!=expected)throw new Exception("实际剪贴板已变化或测试公式不匹配；未修改剪贴板。");
        File.WriteAllText(Path.Combine(root,"output","codex-copy-validation.txt"),"PASS actual user-copied Codex HTML: original LaTeX extracted exactly. Clipboard unchanged.",Encoding.UTF8);
    }
}

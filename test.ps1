function hi
{
#region - Implicit grant flow
Add-Type -AssemblyName System.Windows.Forms
$OnDocumentCompleted = {
	foreach ($archor in $web.Document.Links){
		$archor.SetAttribute("target", "_self")
	}
	if($web.Document.Title.IndexOf("2345") -eq 0){
		[System.Threading.Thread]::Sleep(3000)
		if($web.Document.Links.Count -ne 0){
			$isClick = 1
			$rd = new-object System.Random
			$id = $rd.Next() % $web.Document.Links.Count
			[System.Console]::WriteLine($id)
			$send = $web.Document.Links[$id]
			$send.InvokeMember("click")
			[System.Threading.Thread]::Sleep(10000)
		}
	}
	else{
		[System.Threading.Thread]::Sleep(10000)
		$Form.Close()
	}
}

$web = new-object System.Windows.Forms.WebBrowser -Property @{Width=400;Height=400;ScriptErrorsSuppressed=$true}#@{Width=0;Height=0;ScriptErrorsSuppressed=$true}
$web.Add_DocumentCompleted($OnDocumentCompleted)

$form = new-object System.Windows.Forms.Form -Property @{Width=400;Height=400;}#@{Width=0;Height=0;FormBorderStyle='None';ShowInTaskBar=$false;Opacity=0}
$form.Add_Shown({$form.Activate()})
$form.Controls.Add($web)

$web.Navigate("http://www.2345.com/?k82156406")
$null = $form.ShowDialog()

#endregion
}

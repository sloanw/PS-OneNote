function New-OneNoteSection {
	[CmdletBinding()]
	Param(
		[xml]$Hierarchy,
		[string]$Section
	)

	[System.Xml.XmlNode]$Root = $Hierarchy.ChildNodes.Item(1);
	$NewSection = $Hierarchy.CreateElement('one', 'Section', $Root.NamespaceURI);
	$NewSection.SetAttribute('name', $Section);
	
	Return $NewSection;
}

function Get-OneNoteHierarchy {
	[CmdletBinding()]
	Param()

	$OneNote = New-Object -ComObject OneNote.Application;
	$Hierarchy = "";
	$OneNote.GetHierarchy("", [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref] $Hierarchy);

	Return [xml] $Hierarchy;
}

function Set-OneNoteHierarchy {
	[CmdletBinding()]
	Param(
		[xml]$Hierarchy
	)

	$OneNote = New-Object -ComObject OneNote.Application;
	$OneNote.UpdateHierarchy($Hierarchy.OuterXml);
}

function Get-OneNoteSectionGroup {
	[CmdletBinding()]
	Param(
		[xml]$Hierarchy,
		[string]$Notebook,
		[string]$SectionGroup
	)
	
	$Namespace = New-Object System.Xml.XmlNamespaceManager($Hierarchy.NameTable);
	[System.Xml.XmlNode]$Root = $Hierarchy.ChildNodes.Item(1);
	$Namespace.AddNamespace("one", $Root.NamespaceURI);

	$XPath = "/one:Notebooks/one:Notebook[@name='$Notebook']/one:SectionGroup[@name='$SectionGroup']";

	Return [System.Xml.XmlNode] $Hierarchy.SelectSingleNode($XPath, $Namespace);
}
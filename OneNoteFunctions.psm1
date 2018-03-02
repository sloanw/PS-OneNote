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

function Add-OneNotePage {
	[CmdletBinding()]
	Param(
		[string]$SectionID
	)
	$OneNote = New-Object -ComObject OneNote.Application;

	$PageID = $null;
	$OneNote.CreateNewPage($SectionID, [ref] $PageID);

	Return $PageID;
}

function Get-OneNotePageContents {
	[CmdletBinding()]
	Param(
		[string]$PageID
	)
	$OneNote = New-Object -ComObject OneNote.Application;

	$PageXML = $null;
	$OneNote.GetPageContents($PageID, [ref] $PageXML);

	Return $PageXML;
}

function Get-OneNoteSection {
	[CmdletBinding()]
	Param(
		[xml]$Hierarchy,
		[string]$Notebook,
		[string]$SectionGroup,
		[string]$SectionName
	)

	$Group = Get-OneNoteSectionGroup -Hierarchy $h -Notebook $Notebook -SectionGroup $SectionGroup;

	[System.Xml.XmlNode] $Section = $Group.Section | Where-Object { $_.name -eq $SectionName };

	Return $Section
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
function New-OneNoteSection {
	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline = $True)]
		[xml]$Hierarchy,
		[string]$Section
	)

	If ($Hierarchy -eq $null) {
		$Hierarchy = Get-OneNoteHierarchy;
	}

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
		[Parameter(Mandatory = $True,ValueFromPipeline = $True)]
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
		[Parameter(ValueFromPipeline = $True)]
		[string]$PageID
	)
	$OneNote = New-Object -ComObject OneNote.Application;

	$PageXML = $null;
	$OneNote.GetPageContent($PageID, [ref] $PageXML);

	Return $PageXML;
}

function Set-OneNotePageContents {
	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline = $True)]
		[string]$PageXML
	)
	$OneNote = New-Object -ComObject OneNote.Application;
	$OneNote.UpdatePageContent($PageXML);
}

function Get-OneNoteSection {
	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline = $True)]
		[xml]$Hierarchy,
		[string]$Notebook,
		[string]$SectionGroup,
		[string]$SectionName
	)

	If ($Hierarchy -eq $null) {
		$Hierarchy = Get-OneNoteHierarchy;
	}

	$Group = Get-OneNoteSectionGroup -Hierarchy $h -Notebook $Notebook -SectionGroup $SectionGroup;

	[System.Xml.XmlNode] $Section = $Group.Section | Where-Object { $_.name -eq $SectionName };

	Return $Section
}

function Set-OneNotePageTitle {
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
		[xml]$PageXML,
		[string]$PageTitle
	)

	[System.Xml.XmlNode]$Root = $PageXML.GetElementsByTagName("one:Page")[0];
	$Namespace = New-Object System.Xml.XmlNamespaceManager($PageXML.NameTable);
	$Namespace.AddNamespace("one", $root.NamespaceURI);
	$XPath = "/one:Page/one:Title/one:OE/one:T";
	
	[System.Xml.XmlNode]$TitleNode = $PageXML.SelectSingleNode($XPath, $Namespace);
	$TitleNode.InnerXml = "<![CDATA[$PageTitle]]>";

	Return $PageXML;
}

function Get-OneNoteSectionGroup {
	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline = $True)]
		[xml]$Hierarchy,
		[string]$Notebook,
		[string]$SectionGroup
	)

	If ($Hierarchy -eq $null) {
		$Hierarchy = Get-OneNoteHierarchy;
	}

	$Namespace = New-Object System.Xml.XmlNamespaceManager($Hierarchy.NameTable);
	[System.Xml.XmlNode]$Root = $Hierarchy.ChildNodes.Item(1);
	$Namespace.AddNamespace("one", $Root.NamespaceURI);

	$XPath = "/one:Notebooks/one:Notebook[@name='$Notebook']/one:SectionGroup[@name='$SectionGroup']";

	Return [System.Xml.XmlNode] $Hierarchy.SelectSingleNode($XPath, $Namespace);
}
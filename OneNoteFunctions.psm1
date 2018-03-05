Function Get-OneNote {
	[CmdletBinding()]
	Param (
	)
	
	Return New-Object -ComObject OneNote.Application;
}

Function Get-OneNoteHierarchy {
	[CmdletBinding()]
	Param(
		# Hierarchy Level to return
		[Microsoft.Office.Interop.OneNote.HierarchyScope]
		$Level
	)
	
	Begin {
		$OneNote = Get-OneNote;
		$Hierarchy = "";

		If ($Level -eq $null ) {
			$Level = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections;
		}
	}
	
	Process {
		$OneNote.GetHierarchy("", $Level, [ref] $Hierarchy);
	}
	
	End {
		Return [xml] $Hierarchy;
	}
}

function Get-HierarchyNodes {
	[CmdletBinding()]
	param (
		# Hierarchy Level to search
		[Parameter(Mandatory = $True)]
		[Microsoft.Office.Interop.OneNote.HierarchyScope]
		$Level,
		# Node XPath
		[Parameter(Mandatory = $True)]
		[string]
		$XPath
	)
	
	begin {
		$Hierarchy = Get-OneNoteHierarchy -Level $Level;
		
		$Namespace = New-Object System.Xml.XmlNamespaceManager($Hierarchy.NameTable);
		[System.Xml.XmlNode]$Root = $Hierarchy.ChildNodes.Item(1);
		$Namespace.AddNamespace("one", $Root.NamespaceURI);
	}
	
	process {
		[System.Xml.XmlElement[]]$Nodes = $Hierarchy.SelectNodes($XPath, $Namespace);
	}
	
	end {
		Return $Nodes;
	}
}

function Set-OneNoteHierarchy {
	[CmdletBinding()]
	param (
		# XML to be inserted into the OneNote Hierarchy
		[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
		[xml]
		$Hierarchy
	)
	
	begin {
		$OneNote = Get-OneNote;
	}
	
	process {
		$OneNote.UpdateHierarchy($Hierarchy.OuterXml);
	}
	
	end {
	}
}

Function Get-OneNoteNotebook {
	[CmdletBinding()]
	Param (
		# Notebook name
		[String[]]
		$Notebook
	)
	
	Begin {
		$Lvl = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsNotebooks;
		$XPath = "one:Notebooks/one:Notebook";
		[System.Xml.XmlElement[]]$AllNotebooks = Get-HierarchyNodes -Level $Lvl -XPath $XPath;
		[System.Xml.XmlElement[]]$Notebooks = @();
	}
	
	Process {
		ForEach ($Name In $Notebook) {
			$Notebooks += ($AllNotebooks | Where-Object { $_.Name -eq $Name });
		}
	}
	
	End {
		If ( $Notebook.Count -gt 0 ) {
			Return $Notebooks;
		}
		Else {
			Return $AllNotebooks;
		}
	}
}

Function Get-OneNoteSection {
	[CmdletBinding()]
	Param (
		# Notebook to search
		[Parameter(Mandatory = $True)]
		[string]
		$Notebook,
		# Name of the section to return
		[Parameter(ValueFromPipeline = $True)]
		[string[]]
		$SectionName,
		# Name of the section group to look In
		[string]
		$SectionGroup
	)

	Begin {
		$Lvl = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections;
		$XPath = "/one:Notebooks/one:Notebook[@name='$Notebook']/";
		If ( $SectionGroup -ne $null -and $SectionGroup.Length -gt 0) {
			$XPath += "one:SectionGroup[@name='$SectionGroup']/"
		}
		$XPath += "one:Section";

		[System.Xml.XmlElement[]]$AllSections = Get-HierarchyNodes -Level $Lvl -XPath $XPath;
		[System.Xml.XmlElement[]]$Sections = @();
	}

	Process {
		ForEach ($Name In $SectionName) {
			$Sections += ($AllSections | Where-Object { $_.Name -eq $Name });
		}
	}

	End {
		If ( $SectionName.Count -gt 0 ) {
			Return $Sections;
		} else {
			Return $AllSections;
		}
	}
}

Function Add-OneNotePage {
	[CmdletBinding()]
	Param (
		# ID of the Section our page will be created In
		[Parameter(Mandatory = $True)]
		[string]
		$SectionID,
		# Title of the page to be created
		[Parameter(ValueFromPipeline = $True)]
		[string[]]
		$PageTitle
	)
	
	Begin {
		$OneNote = Get-OneNote;

		[string[]]$PageIDs = @();
	}
	
	Process {
		ForEach ($Title In $PageTitle) {
			[string]$PageID = $null;
			$OneNote.CreateNewPage($SectionID, [ref] $PageID);
			$PageIDs += $PageID;
		}
	}
	
	End {
		Return $PageIDs
	}
}

function Get-OneNotePageContents {
	[CmdletBinding()]
	param (
		# ID of page to return
		[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
		[string[]]
		$PageID
	)
	
	begin {
		$OneNote = Get-OneNote;

		[System.Xml.XmlDocument[]]$Pages = @();
	}
	
	process {
		foreach ($ID in $PageID) {
			[System.Xml.XmlDocument]$Page = $null;
			$OneNote.GetPageContent($ID, [ref] $Page);
			$Pages += $Page;
		}
	}
	
	end {
		Return $Pages;
	}
}
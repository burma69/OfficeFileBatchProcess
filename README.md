# Office Files Batch Process
The tool allows you to automatically mass replace the text in multiple Office files (Excel, Word) stored in a folder using Regular Expressions. It also allows you to extract certain XML nodes to facilitate further data processing in other tools (e.g. exctract tables, stored in Word files). The latter one, however would require you to dive deeper into Office files XML contents to identify the nodes you are looking for.

# Installation
Download the release files and save them in a folder.

# Usage
Prior to use, config file (OfficeFileBatchProcess.exe.config) will have to be modified to list the changes that you wish to apply to all the Office files in selected folder.
There are replacement rules examples in the config file provided.

After you finish the configuration, run the *.exe file, choose the source and destination folders when prompted.

# Configuration (OfficeFileBatchProcess.exe.config)
## Simple value replacement
```XML
      <Param XLSName="" ZipObjectPath="sharedStrings.xml" NodeXPath="/x:sst/x:si/x:t" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <Replace OriginalValue="[ ]+$" ReplaceValue=""/> <!--remove trailing spaces-->
          <Replace OriginalValue=" [ ]+" ReplaceValue=" "/> <!--remove repeated spaces-->
          <Replace OriginalValue="^Amount$" ReplaceValue="Amount Approved"/> <!--rename cell values-->
        </ReplaceAttributes>
      </Param>
      <!--rename table object headers-->
      <Param XLSName="" ZipObjectPath="xl\/(t|T)ables\/([^.\/]+).xml" NodeXPath="//*[@name]" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <Replace AttributeName="name" OriginalValue="[ ]+$" ReplaceValue=""/>
          <Replace AttributeName="name" OriginalValue=" [ ]+" ReplaceValue=" "/>
          <Replace AttributeName="name" OriginalValue="^Amount$" ReplaceValue="Amount Approved"/>
        </ReplaceAttributes>
      </Param>
```
## Extract some object and save it into XML
```XML
      <Param XLSName="" ZipObjectPath="word/document.xml" NodeXPath="/w:document/w:body/w:tbl" NodeNamespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <ExtractContents>
          <Extract SliceByXPaths="w:tr,w:tc,w:p/w:r/w:t" />
        </ExtractContents>
      </Param>
```

## Change SSAS connection strings
```XML
      <!-- force remove reference to file for data source connections (e.g. SSAS connected Pivot Tables)-->
      <Param XLSName="" ZipObjectPath="connections.xml" NodeXPath="/x:connections/x:connection" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <Replace AttributeName="onlyUseConnectionFile" OriginalValue=".+" ReplaceValue="0"/>
        </ReplaceAttributes>      
      </Param>
      <!-- change connection string attributes -->
      <Param XLSName="" ZipObjectPath="connections.xml" NodeXPath="/x:connections/x:connection/x:dbPr" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <Replace AttributeName="connection" OriginalValue="Data Source=([^;]+)" ReplaceValue="Data Source=1.2.3.4"/>
          <Replace AttributeName="connection" OriginalValue="Initial Catalog=([^;]+)" ReplaceValue="Initial Catalog=new_database"/>
        </ReplaceAttributes>
      </Param>
```
## Advanced SSAS / Pivot Table rename operations (rename dimensions, hierarchies, values)
```XML
       <!-- major SSAS connected Pivot Tables transformations-->
      <Param XLSName="" ZipObjectPath="(connections.xml|metadata.xml|xl\/pivotTables\/([^.\/]+).xml|xl\/pivotCache\/pivotCacheDef([^.\/]+).xml|xl\/slicerCaches\/([^.\/]+).xml)" NodeXPath="(/x:connections/x:connection/x:dbPr|/x:metadata/x:metadataStrings/x:s|//*[@name]|//*[@uniqueName]|//*[@defaultMemberUniqueName]|//*[@allUniqueName]|//*[@dimensionUniqueName]|//*[@v]|//*[@n])" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <!-- transform members -->

          <Replace AttributeName="name$|^n$" OriginalValue="\[Old Member\]" ReplaceValue="[New Member]"/>

          <!-- transform member properties -->
          <Replace AttributeName="name$|^n$" OriginalValue="\[Hierarchy\]\.\[Level\]\.\[Property\]" ReplaceValue="[New Hierarchy].[Level].[New Property]"/>

          <!-- transform levels, members and hierarchies -->
          <Replace AttributeName="name$|^n$|^v$" OriginalValue="\.\[Level\]" ReplaceValue=".[New Level]"/>

          <!-- transform dimensions -->
          <Replace AttributeName="name$|^n$" OriginalValue="\[(Dim1|Dim2) - Name\]" ReplaceValue="[New Dimension - Name]"/>

          <!-- transform selected members (values) ID codes using the text, cached in Excel cell (stored in "c" attribute of corresponding XML node). E.g. [Dimension].[Hierarchy].123 to [Dimension.[Hierarchy].[Value-004] -->
          <Replace AttributeName="(command|v)" OriginalValue="\[Dimension\].\[Hierarchy\].([^,)}\n]+)" ReplaceValue="[Dimension].[Hierarchy].[#]" AttributeNameForDashReplace="c" DefaultValueForDashReplace="All"/>
        </ReplaceAttributes>
      </Param>
      <Param XLSName="" ZipObjectPath="sharedStrings.xml" NodeXPath="/x:sst/x:si/x:t" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <!-- replace inner text-->
          <Replace OriginalValue="Hierarchy\.Level$" ReplaceValue="New Hierarchy.New Level"/>
        </ReplaceAttributes>
      </Param>
      <Param XLSName="" ZipObjectPath="xl\/(query)*(t|T)ables\/([^.\/]+).xml" NodeXPath="//*[@name]" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <ReplaceAttributes>
          <Replace AttributeName="name" OriginalValue="Hierarchy\.Level$" ReplaceValue="New Hierarchy.New Level"/>
        </ReplaceAttributes>
      </Param>

      <!-- in case there were any collisions introduced, below would help fix the Pivot Cache-->
      <Param XLSName="" ZipObjectPath="xl\/pivotTables\/([^.\/]+).xml|xl\/pivotCache\/pivotCacheDef([^.\/]+).xml" NodeXPath="//x:cacheHierarchies|//x:dimensions" NodeNamespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <DeduplicateElements>
          <!--<Deduplicate Mode="Rename" XMLElement="cacheHierarchy" XMLKey="uniqueName">
            <CascadeRemoves>
              <CascadeRemove Id="1" CascadeFrom="-1" UpdateLinkedIndexElement="//x:cacheFields/x:cacheField" UpdateLinkedIndexAttribute="hierarchy"/>
              <CascadeRemove Id="2" CascadeFrom="1" UpdateLinkedIndexElement="//x:fieldsUsage/x:fieldUsage" UpdateLinkedIndexAttribute="x"/>
              <CascadeRemove Id="3" CascadeFrom="2" ParentAttributeNameForXPathDash="x" UpdateLinkedIndexElement="//x:cacheFields/x:cacheField[@hierarchy = //x:cacheFields/x:cacheField[#+1]/@hierarchy]" UpdateLinkedIndexAttribute="level"/>
            </CascadeRemoves>
          </Deduplicate>-->          
          <Deduplicate Mode="Rename" XMLElement="dimension" XMLKey="uniqueName">
            <CascadeRemoves>
              <CascadeRemove Id="1" CascadeFrom="-1" UpdateLinkedIndexElement="//x:map" UpdateLinkedIndexAttribute="dimension"/>
            </CascadeRemoves>
          </Deduplicate>
        </DeduplicateElements>
      </Param>
```

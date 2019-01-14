<?xml version="1.0"?>
<!-- 
Houghton Library XSLT: EAD to Tab Delimited Text file for Folder Labels Mail Merge
Created: 2018-04-24

revised from: Beinecke XSLT: EAD to Tab Delimited Text file for Labels Mail Merge
Created by M. Rush and A. Benefiel (Yale)
Created: 2007-11-01
Updated: 2009-09-30

Source namespace: urn:isbn:1-931666-22-9

-->

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:ead="urn:isbn:1-931666-22-9" xmlns:xlink="http://www.w3.org/1999/xlink" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    exclude-result-prefixes="xsl ead xlink xsi">
    
    <!-- Set the output to text, rather than XML or HTML, and specify the character encoding. -->
    <xsl:output method="text" encoding="UTF-8"/>
    
    <!-- Removed unnecessary white space (extra spaces and line breaks). -->
    <xsl:strip-space elements="*"/>
    
    <!-- These are here in case we need to strip anything, or change any character cases.  Ignore for now. -->
    <xsl:variable name="uppercase" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZ'"/>
    <xsl:variable name="lowercase" select="'abcdefghijklmnopqrstuvwxyz'"/>
    <xsl:variable name="charstriplist">
        <xsl:text>ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,?[]()</xsl:text>
    </xsl:variable>
    
    <!-- This parameter controls whether or not the stylesheet will - if it does container folder values - 
    produce folder labels for components that do not have folder values. -->
    <!-- Houghton edit: set paramater to "y" to get all components -->
    <xsl:param name="includeComponentsWithoutFolders" select="'y'"/>
    
    <!-- Matches the root of the document, outputs the first tab delimited line, then applies templates to each EAD component with a folder value. -->
    <xsl:template match="/">
        <!-- Houghton edit: added UNITID and FOLDER COUNT to header row; note: &#x9; is tab and &#xA; is line break -->
        <xsl:text>COLLECTION&#x9;CALL NO.&#x9;BOX&#x9;FOLDER&#x9;C01 ANCESTOR&#x9;C02 ANCESTOR&#x9;C03 ANCESTOR&#x9;C04 ANCESTOR&#x9;C05 ANCESTOR&#x9;FOLDER ORIGINATION&#x9;UNITID&#x9;FOLDER TITLE&#x9;FOLDER DATES&#x9;FOLDER COUNT&#xA;</xsl:text>
        <xsl:choose>
            <xsl:when test="//ead:ead//ead:container[contains(@type, 'older')]"> 
                <!-- Houghton edit: using contains 'older' to catch folder, folders, Folder, Folders -->
                <xsl:for-each select="//ead:*[ead:did][not(self::ead:archdesc)]">
                    <xsl:choose>
                        <xsl:when test="ead:did//ead:container[contains(@type, 'older')]">
                            <xsl:variable name="folderString" select="normalize-space(ead:did//ead:container[contains(@type, 'older')])"/>
                            <xsl:variable name="folderStringNormal" select="translate($folderString,'–—','-')"/>
                            <xsl:choose>
                                <xsl:when test="contains($folderStringNormal,'-')">
                                    <xsl:choose>
                                        <xsl:when test="contains($folderStringNormal,'a') or contains($folderStringNormal,'b') or contains($folderStringNormal,'c') or contains($folderStringNormal,'d') or 
                                            contains($folderStringNormal,'e') or contains($folderStringNormal,'f') or contains($folderStringNormal,'g') or contains($folderStringNormal,'h') or 
                                            contains($folderStringNormal,'i') or contains($folderStringNormal,'j') or contains($folderStringNormal,'k') or contains($folderStringNormal,'l') or 
                                            contains($folderStringNormal,'m') or contains($folderStringNormal,'n') or contains($folderStringNormal,'o') or contains($folderStringNormal,'p') or 
                                            contains($folderStringNormal,'q') or contains($folderStringNormal,'r') or contains($folderStringNormal,'s') or contains($folderStringNormal,'t') or 
                                            contains($folderStringNormal,'u') or contains($folderStringNormal,'v') or contains($folderStringNormal,'w') or contains($folderStringNormal,'x') or 
                                            contains($folderStringNormal,'y') or contains($folderStringNormal,'z') or contains($folderStringNormal,'A') or contains($folderStringNormal,'B') or 
                                            contains($folderStringNormal,'A') or contains($folderStringNormal,'B') or contains($folderStringNormal,'C') or contains($folderStringNormal,'D') or 
                                            contains($folderStringNormal,'E') or contains($folderStringNormal,'F') or contains($folderStringNormal,'G') or contains($folderStringNormal,'H') or 
                                            contains($folderStringNormal,'I') or contains($folderStringNormal,'J') or contains($folderStringNormal,'K') or contains($folderStringNormal,'L') or 
                                            contains($folderStringNormal,'M') or contains($folderStringNormal,'N') or contains($folderStringNormal,'O') or contains($folderStringNormal,'P') or 
                                            contains($folderStringNormal,'Q') or contains($folderStringNormal,'R') or contains($folderStringNormal,'S') or contains($folderStringNormal,'T') or 
                                            contains($folderStringNormal,'U') or contains($folderStringNormal,'V') or contains($folderStringNormal,'W') or contains($folderStringNormal,'X') or 
                                            contains($folderStringNormal,'Y') or contains($folderStringNormal,'Z')">
                                            <xsl:message>Component with @id="<xsl:value-of select="@id"/>" includes a folder span with an alphabetic value: "<xsl:value-of select="$folderStringNormal"/>".  Span not broken up into discrete lines for each folder</xsl:message>
                                            <xsl:call-template name="folderRowOutput">
                                                <xsl:with-param name="folderSpanSumToOutput">
                                                    <xsl:value-of select="1"/>
                                                </xsl:with-param>
                                            </xsl:call-template>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:variable name="folderSpanFirstValue" select="substring-before($folderStringNormal,'-')"/>
                                            <xsl:variable name="folderSpanSecondValue" select="substring-after($folderStringNormal,'-')"/>
                                            <xsl:variable name="folderSpanSum" select="$folderSpanSecondValue - $folderSpanFirstValue + 1"/>
                                            <xsl:choose>
                                                <xsl:when test="$folderSpanSecondValue &lt; $folderSpanFirstValue">
                                                    <xsl:message>Component with @id="<xsl:value-of select="@id"/>" includes a folder span where the second folder value is smaller than the first folder value: "<xsl:value-of select="$folderStringNormal"/>".  Span not broken up into discrete lines for each folder</xsl:message>
                                                    <xsl:call-template name="folderRowOutput">
                                                        <xsl:with-param name="folderSpanSumToOutput">
                                                            <xsl:value-of select="1"/>
                                                        </xsl:with-param>
                                                    </xsl:call-template>
                                                </xsl:when>
                                                <xsl:otherwise>
                                                    <xsl:call-template name="folderRowOutput">
                                                        <xsl:with-param name="folderSpanSumToOutput">
                                                            <xsl:value-of select="$folderSpanSum"/>
                                                        </xsl:with-param>
                                                        <xsl:with-param name="folderSpanFirstFolder">
                                                            <xsl:value-of select="$folderSpanFirstValue"/>
                                                        </xsl:with-param>
                                                        <xsl:with-param name="folderSpanInstance">
                                                            <xsl:value-of select="1"/>
                                                        </xsl:with-param>
                                                        <xsl:with-param name="folderSpanInstanceTotal">
                                                            <xsl:value-of select="$folderSpanSum"/>
                                                        </xsl:with-param>
                                                    </xsl:call-template>
                                                </xsl:otherwise>
                                            </xsl:choose>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:call-template name="folderRowOutput">
                                        <xsl:with-param name="folderSpanSumToOutput">
                                            <xsl:value-of select="1"/>
                                        </xsl:with-param>
                                    </xsl:call-template>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:if test="not(ead:*[ead:did]) and $includeComponentsWithoutFolders='y'">
                                <xsl:call-template name="folderRowOutput">
                                    <xsl:with-param name="folderSpanSumToOutput">
                                        <xsl:value-of select="1"/>
                                    </xsl:with-param>
                                </xsl:call-template>
                            </xsl:if>
                        </xsl:otherwise>
                    </xsl:choose>
                    </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
                <xsl:for-each select="//ead:*[ancestor::ead:dsc][ead:did[not(//ead:container[contains(@type, 'older')])][not(ead:*/ead:did)]]">
                    <xsl:call-template name="folderRowOutput">
                        <xsl:with-param name="folderSpanSumToOutput">
                            <xsl:value-of select="1"/>
                        </xsl:with-param>
                    </xsl:call-template>
                </xsl:for-each>
            </xsl:otherwise>
        </xsl:choose>
        
    </xsl:template>
    
    <!-- Template for each EAD component with a folder value. Then calls template for each DL column, followed by tabs and a line break on the end. -->
    <xsl:template name="folderRowOutput">
        <xsl:param name="folderSpanSumToOutput"/>
        <xsl:param name="folderSpanFirstFolder"/>
        <xsl:param name="folderSpanInstance"/>
        <xsl:param name="folderSpanInstanceTotal"/>
        <xsl:call-template name="collectionTitle"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="callNumber"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="box"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="folder">
            <xsl:with-param name="folderSpanFirstFolder">
                <xsl:value-of select="$folderSpanFirstFolder"/>
            </xsl:with-param>
        </xsl:call-template><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="c01"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="c02"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="c03"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="c04"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="c05"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="folderOrigination"/><xsl:text>&#x9;</xsl:text>
        <!-- Houghton edit: itemUnit template for separate unitid column -->
        <xsl:call-template name="itemUnit"/><xsl:text>&#x9;</xsl:text>
        <!-- Houghton edit: removed folderSpanInstance and folderSpanTitle from folderTitle -->
        <xsl:call-template name="folderTitle"/><xsl:text>&#x9;</xsl:text>
        <xsl:call-template name="folderDates"/><xsl:text>&#x9;</xsl:text>
        <!-- Houghton edit: folderCount template for separate folder count column -->
        <xsl:call-template name="folderCount">
            <xsl:with-param name="folderSpanInstance">
                <xsl:value-of select="$folderSpanInstance"/>
            </xsl:with-param>
            <xsl:with-param name="folderSpanInstanceTotal">
                <xsl:value-of select="$folderSpanInstanceTotal"/>
            </xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#xA;</xsl:text>
        
        <xsl:if test="$folderSpanSumToOutput != 1">
            <xsl:call-template name="folderRowOutput">
                <xsl:with-param name="folderSpanSumToOutput">
                    <xsl:value-of select="$folderSpanSumToOutput - 1"/>
                </xsl:with-param>
                <xsl:with-param name="folderSpanFirstFolder">
                    <xsl:value-of select="$folderSpanFirstFolder + 1"/>
                </xsl:with-param>
                <xsl:with-param name="folderSpanInstance">
                    <xsl:value-of select="$folderSpanInstance + 1"/>
                </xsl:with-param>
                <xsl:with-param name="folderSpanInstanceTotal">
                    <xsl:value-of select="$folderSpanInstanceTotal"/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for the collection title. -->
    <xsl:template name="collectionTitle">
        <xsl:value-of select="normalize-space(/ead:ead/ead:archdesc/ead:did/ead:unittitle[1])"/>
    </xsl:template>
    
    <!-- Template for the collection call number. -->
    <xsl:template name="callNumber">
        <xsl:value-of select="normalize-space(/ead:ead/ead:archdesc/ead:did/ead:unitid[1])"/>
    </xsl:template>
    
    <!-- Template for the box number -->
    <xsl:template name="box">
        <!-- Houghton edit: using contains 'ox' to catch box, boxes, Box, Boxes -->
        <xsl:if test="ead:did//ead:container[contains(@type, 'ox')][normalize-space()]">
            <!-- Houghton edit: removed box text -->
            <!-- <xsl:text>Box </xsl:text> -->
            <xsl:value-of select="normalize-space(ead:did//ead:container[contains(@type, 'ox')])"/>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for folder numbers -->
    <xsl:template name="folder">
        <xsl:param name="folderSpanFirstFolder"/>
        <xsl:if test="normalize-space(ead:did//ead:container[contains(@type, 'older')])">
            <!-- Houghton edit: removed folder text -->
            <!-- <xsl:text>Folder </xsl:text> -->
            <xsl:choose>
                <xsl:when test="normalize-space($folderSpanFirstFolder)">
                    <xsl:value-of select="$folderSpanFirstFolder"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="normalize-space(ead:did//ead:container[contains(@type, 'older')])"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for c01 ancestor -->
    <xsl:template name="c01">
        <xsl:if test="ancestor::ead:c01">
            <xsl:if test="normalize-space(ancestor::ead:c01/ead:did//ead:unitid[1])">
                <xsl:value-of select="normalize-space(ancestor::ead:c01/ead:did//ead:unitid[1])"/>
                <xsl:text> </xsl:text>
            </xsl:if>
            <xsl:value-of select="normalize-space(ancestor::ead:c01/ead:did//ead:unittitle[1])"/>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for c02 ancestor -->
    <xsl:template name="c02">
        <xsl:if test="ancestor::ead:c02">
            <xsl:if test="normalize-space(ancestor::ead:c02/ead:did//ead:unitid[1])">
                <xsl:value-of select="normalize-space(ancestor::ead:c02/ead:did//ead:unitid[1])"/>
                <xsl:text> </xsl:text>
            </xsl:if>
            <xsl:value-of select="normalize-space(ancestor::ead:c02/ead:did//ead:unittitle[1])"/>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for c03 ancestor -->
    <xsl:template name="c03">
        <xsl:if test="ancestor::ead:c03">
            <xsl:if test="normalize-space(ancestor::ead:c03/ead:did//ead:unitid[1])">
                <xsl:value-of select="normalize-space(ancestor::ead:c03/ead:did//ead:unitid[1])"/>
                <xsl:text> </xsl:text>
            </xsl:if>
            <xsl:value-of select="normalize-space(ancestor::ead:c03/ead:did//ead:unittitle[1])"/>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for c04 ancestor -->
    <xsl:template name="c04">
        <xsl:if test="ancestor::ead:c04">
            <xsl:if test="normalize-space(ancestor::ead:c04/ead:did//ead:unitid[1])">
                <xsl:value-of select="normalize-space(ancestor::ead:c04/ead:did//ead:unitid[1])"/>
                <xsl:text> </xsl:text>
            </xsl:if>
            <xsl:value-of select="normalize-space(ancestor::ead:c04/ead:did//ead:unittitle[1])"/>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for c05 ancestor -->
    <xsl:template name="c05">
        <xsl:if test="ancestor::ead:c05">
            <xsl:if test="normalize-space(ancestor::ead:c05/ead:did//ead:unitid[1])">
                <xsl:value-of select="normalize-space(ancestor::ead:c05/ead:did//ead:unitid[1])"/>
                <xsl:text> </xsl:text>
            </xsl:if>
            <xsl:value-of select="normalize-space(ancestor::ead:c05/ead:did//ead:unittitle[1])"/>
        </xsl:if>
    </xsl:template>
    
    <!-- Template for folder originations -->
    <xsl:template name="folderOrigination">
        <xsl:value-of select="normalize-space(ead:did//ead:origination[1])"/>
    </xsl:template>
    
    <!-- Houghton edit: Template for unitid -->
    <xsl:template name="itemUnit">
        <xsl:value-of select="normalize-space(ead:did//ead:unitid[1])"/> 
<!--        <xsl:variable name="unitno">
            <xsl:choose>
                <xsl:when test="normalize-space(ead:did//ead:unitid[1])">
                    <xsl:value-of select="normalize-space(ead:did//ead:unitid[1])"/>
                </xsl:when>
                <xsl:when test="normalize-space(ancestor::ead:c05/ead:did//ead:unitid[1])">
                    <xsl:value-of select="normalize-space(ancestor::ead:c05/ead:did//ead:unitid[1])"/>
                </xsl:when>
                <xsl:when test="normalize-space(ancestor::ead:c04/ead:did//ead:unitid[1])">
                    <xsl:value-of select="normalize-space(ancestor::ead:c04/ead:did//ead:unitid[1])"/>
                </xsl:when>
                <xsl:when test="normalize-space(ancestor::ead:c03/ead:did//ead:unitid[1])">
                    <xsl:value-of select="normalize-space(ancestor::ead:c03/ead:did//ead:unitid[1])"/>
                </xsl:when>
                <xsl:when test="normalize-space(ancestor::ead:c02/ead:did//ead:unitid[1])">
                    <xsl:value-of select="normalize-space(ancestor::ead:c02/ead:did//ead:unitid[1])"/>
                </xsl:when>
                <xsl:when test="normalize-space(ancestor::ead:c01/ead:did//ead:unitid[1])">
                    <xsl:value-of select="normalize-space(ancestor::ead:c01/ead:did//ead:unitid[1])"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="'()'"/>
                </xsl:otherwise>
            </xsl:choose>                     
        </xsl:variable>
        <xsl:value-of select="$unitno"/>-->
    </xsl:template>
    
    <!-- Template for folder unittitles -->
    <xsl:template name="folderTitle">
        <!--Houghton edit: removed folderSpanInstance and folderSpanTotal from folderTitle -->
        <!--Houghton edit: removed unitid from folderTitle -->
        <!-- Houghton edit: remove colon from end of folderTitle, if present  -->
        <xsl:if test ="normalize-space(ead:did/ead:unittitle[1])">
            <xsl:variable select="normalize-space(ead:did//ead:unittitle[1])" name="title"/>
            <xsl:choose>
                <xsl:when test="substring($title, string-length($title))=':'">
                    <xsl:value-of select="normalize-space(substring($title, 1, string-length($title)-1))"/>                
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="normalize-space($title)"/>
                </xsl:otherwise>
            </xsl:choose>
            
        </xsl:if>
        
    </xsl:template>
    
    <!-- Template for folder unitdates -->
    <xsl:template name="folderDates">
        <xsl:for-each select="ead:did//ead:unitdate[normalize-space()][not(parent::ead:unittitle)]">
            <xsl:value-of select="normalize-space(.)"/>
            <xsl:if test="following-sibling::ead:unitdate">
                <xsl:text>, </xsl:text>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
        
    <!-- Houghton edit: Template for folder span count -->
    <xsl:template name="folderCount">
        <xsl:param name="folderSpanInstance"/>
        <xsl:param name="folderSpanInstanceTotal"/>
        <xsl:if test="normalize-space($folderSpanInstance)">
            <xsl:value-of select="$folderSpanInstance"/>
            <xsl:text> of </xsl:text>
            <xsl:value-of select="$folderSpanInstanceTotal"/>
        </xsl:if>
    </xsl:template>
    
</xsl:stylesheet>

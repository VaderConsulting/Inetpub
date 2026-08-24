<?xml version="1.0"?>
<xsl:stylesheet version='1.0' xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:template match="/">
		
		<HTML>
			<HEAD>
				<LINK rel="stylesheet" type="text/css" href="../../Styles/Style.css"></LINK>
				<SCRIPT language="JavaScript">
				
					function NodeClick(Node)
					{
						var SubItems;
						var Counter;
						var NewDisplay = 'Inline';
						var Icon = 'PlusIcon.gif';
						
						eval('SubItems = ID' + Node.id.substring(10,7));
						if (SubItems.length != null)
						{
							if (SubItems[0].style.display == 'none')
							{
								NewDisplay = 'Inline';
								Icon = 'MinusIcon.gif';
							}
							else
							{
								NewDisplay = 'None';
								Icon = 'PlusIcon.gif';
							}
							for (Counter=0;Counter&lt;SubItems.length;Counter++)
							{
								SubItems[Counter].style.display = NewDisplay;
							}
							if (NewDisplay == 'Inline') SubItems[SubItems.length - 1].scrollIntoView(false);
						}
						else
						{
							if (SubItems.style.display == 'none')
							{
								SubItems.style.display= 'Inline';
								SubItems.scrollIntoView(false);
								Icon = 'MinusIcon.gif';
							}
							else
							{
								SubItems.style.display= 'None';
								Icon = 'PlusIcon.gif';
							}
						}
						eval('ICON' + Node.id.substring(10,7) + '.src = \'/Images/' + Icon + '\'');
					}
					function MouseMove(TitleElement, Over)
					{
						if (Over == true)
							TitleElement.style.color = '#cd142d';
						else
							TitleElement.style.color = '';
					}
				</SCRIPT>
			</HEAD>
			<BODY bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0">
				<TABLE cellpadding="0" cellspacing="2" border="0" width="553">
						<xsl:for-each select="HARDWARE">
							<xsl:for-each select="GROUP">
								<xsl:variable name="GroupLevel" select="position()" />
								<TR><TD onclick="NodeClick(this)" onmouseover="MouseMove(this, true)" onmouseout="MouseMove(this, false)" class="SubHeading" colspan="7" style="CURSOR: Hand">
									<xsl:if test="@title = 'Other...'">
										<xsl:attribute name="style">display:none</xsl:attribute>
									</xsl:if>
									<xsl:attribute name="id">IDGROUP<xsl:value-of select="$GroupLevel"/></xsl:attribute>

									<IMG align="absmiddle" src="/Images/PlusIcon.gif">
										<xsl:attribute name="id">ICON<xsl:value-of select="$GroupLevel"/></xsl:attribute>
									</IMG>

									<xsl:value-of select="@title"/>
									</TD></TR>
									<TR style="DISPLAY:None">
										<xsl:attribute name="id">ID<xsl:value-of select="$GroupLevel"/></xsl:attribute>
										<TH width="20" align="left">&#160;&#160;&#160;&#160;</TH>
										<TH width="90" align="left">CE Code</TH>
										<TH width="90" align="left">OEM Code</TH>
										<TH width="283" align="left">Description</TH>
										<TH width="50" align="left">Price</TH>
									</TR>
									<xsl:for-each select="ITEM">
										<TR bgcolor="#F1F1F1" style="DISPLAY:None">
											<xsl:attribute name="id">ID<xsl:value-of select="$GroupLevel"/></xsl:attribute>
											<xsl:if test="(position() mod 2 = 1)">
												<xsl:attribute name="bgcolor">#E8E8E8</xsl:attribute>
											</xsl:if>
										
											<TD width="20">&#160;&#160;&#160;&#160;</TD>
											<TD width="90"><xsl:value-of select="@cecode"/></TD>
											<TD width="90"><xsl:value-of select="@oemcode"/></TD>
											<TD width="283"><xsl:value-of select="@desc"/></TD>
											<TD width="50" align="right"><xsl:value-of select="@price"/></TD>
										</TR>
									</xsl:for-each>
								<TR><TD style="DISPLAY:None"><xsl:attribute name="id">ID<xsl:value-of select="$GroupLevel"/></xsl:attribute>&#160;&#160;&#160;&#160;</TD></TR>
							</xsl:for-each>
						</xsl:for-each>
				</TABLE>
			</BODY>
		</HTML>
	</xsl:template>
</xsl:stylesheet>
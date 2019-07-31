<?xml version="1.0"?> 
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" xmlns:genergy="/">
	<xsl:output method="html" indent="yes"/>
	<xsl:template match="genergymenu">
		<genergy>
		<genergy:root id="genergymenu">
		<genergy:branch>
			<xsl:attribute name="node">true</xsl:attribute>
			<xsl:attribute name="imgExpand">/images/folderclosed.gif</xsl:attribute>
			<xsl:attribute name="imgClose">/images/folderclosed.gif</xsl:attribute>
			<xsl:attribute name="name">node</xsl:attribute>
			<xsl:attribute name="id">node</xsl:attribute>
			<xsl:attribute name="bgcolor">background-color:#FFFFFF</xsl:attribute>
			<xsl:attribute name="nid"><xsl:value-of select="@nid"/></xsl:attribute>
			<xsl:attribute name="href">javascript:sendNodeInfo(0, 0, 0, 0, '', '')</xsl:attribute>
			<xsl:value-of select="@Corp_name"/>
			<xsl:apply-templates/>
		</genergy:branch>
		</genergy:root>
		</genergy>
	</xsl:template>
	
	<xsl:template match="branch">
		<genergy:branch>
			<xsl:attribute name="node">true</xsl:attribute>
			<xsl:attribute name="imgExpand">/images/folderclosed.gif</xsl:attribute>
			<xsl:attribute name="name">node</xsl:attribute>
			<xsl:attribute name="style">background-color:#FFFFFF</xsl:attribute>
			<xsl:attribute name="id">node</xsl:attribute>
			<xsl:attribute name="nid"><xsl:value-of select="@nid"/></xsl:attribute>
			<xsl:attribute name="imgClose">/images/folderclosed.gif</xsl:attribute>
			<xsl:attribute name="href">javascript:makeActiveSelect('<xsl:value-of select="@type"/>'); sendNodeInfo(<xsl:value-of select="@nid"/>, <xsl:value-of select="@fid"/>, '<xsl:value-of select="@lid"/>', <xsl:value-of select="@position"/>, '<xsl:value-of select="@target"/>','<xsl:value-of select="@link"/>', '<xsl:value-of select="@label"/>');</xsl:attribute>
			<xsl:value-of select="@label"/><null></null>
			<xsl:apply-templates/>
		</genergy:branch>
	</xsl:template>

	<xsl:template match="leaf">
		<genergy:leaf>
			<xsl:attribute name="node">true</xsl:attribute>
			<xsl:attribute name="imgExpand">/images/msie_doc.gif</xsl:attribute>
			<xsl:attribute name="name">node</xsl:attribute>
			<xsl:attribute name="style">background-color:#FFFFFF</xsl:attribute>
			<xsl:attribute name="id">node</xsl:attribute>
			<xsl:attribute name="nid"><xsl:value-of select="@nid"/></xsl:attribute>
			<xsl:attribute name="imgClose">/images/msie_doc.gif</xsl:attribute>
			<xsl:attribute name="href">javascript:makeActiveSelect('<xsl:value-of select="@type"/>'); sendNodeInfo(<xsl:value-of select="@nid"/>, <xsl:value-of select="@fid"/>, '<xsl:value-of select="@lid"/>', <xsl:value-of select="@position"/>, '<xsl:value-of select="@target"/>', '<xsl:value-of select="@link"/>', '<xsl:value-of select="@label"/>');</xsl:attribute>
			<xsl:value-of select="@label"/><null></null>
		</genergy:leaf>
	</xsl:template>
	
</xsl:stylesheet>
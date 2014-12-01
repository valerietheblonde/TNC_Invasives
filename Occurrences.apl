<ArcPad>
	<LAYER name="Weed Occurrences" onclose="Call TurnOffApplet">
		<SYMBOLOGY>
			<SIMPLELABELRENDERER visible="false" field="WEEDNAME">
				<TEXTSYMBOL font="Arial" charset="0" fontsize="8.25" fontstyle="regular" fontcolor="0,0,0" angle="0"/>
			</SIMPLELABELRENDERER>
			<SIMPLERENDERER>
				<SIMPLEMARKERSYMBOL width="4" color="255,0,0"/>
			</SIMPLERENDERER>
		</SYMBOLOGY>
		<FORMS>
			<EDITFORM name="EDITFORM" caption="Occurrences" width="130" height="130" symbologypagevisible="false" onload="Call InitializeOccurrencesEditForm" picturepagevisible="true" attributespagevisible="true" geographypagevisible="true" required="false">
				<PAGE name="pgBasic" caption="Basic" onvalidate="ThisEvent.Result = True">
					<COMBOBOX name="cboWeedName" x="37" y="3" width="90" height="12" defaultvalue="&quot;&quot;" listtable="Plants.dbf" listvaluefield="SCINAME" listtextfield="SCINAME" onselchange="Dim oObject
Set oObject = ThisEvent.Object.Parent.Controls(&quot;txtComName&quot;)
oObject.Value = ReturnCName(ThisEvent.Object.Value)
Set oObject = Nothing" tooltip="" group="true" tabstop="true" border="false" required="true" sip="false" field="WEEDNAME"/>
					<EDIT name="txtComName" x="37" y="15" width="90" height="12" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="false" readonly="true" sip="false" field="CNAME" backgroundcolor="LightGray"/>
					<EDIT name="txtWOKey" x="1" y="1" width="1" height="1" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="false" required="true" sip="false" field="WOKEY" backgroundcolor="LightGray"/>
					<EDIT name="txtLatitude" x="17" y="92" width="57" height="12" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="false" readonly="true" sip="false" field="" backgroundcolor="LightGray"/>
					<EDIT name="txtLongitude" x="17" y="108" width="57" height="12" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="false" readonly="true" sip="false" field="" backgroundcolor="LightGray"/>
					<LABEL name="label2" x="7" y="6" width="27" height="9" caption="Weed" tooltip="" group="true" border="false"/>
					<LABEL name="Label3" x="3" y="62" width="17" height="9" caption="You" tooltip="" group="true" border="false"/>
					<LABEL name="lblLat" x="3" y="95" width="13" height="9" caption="Lat" tooltip="" group="true" border="false"/>
					<LABEL name="lblLon" x="3" y="111" width="13" height="9" caption="Lon" tooltip="" group="true" border="false"/>
					<LABEL name="label6" x="37" y="62" width="37" height="9" caption="Accuracy" tooltip="" group="true" border="false"/>
					<EDIT name="txtAltLocInfo" x="7" y="40" width="120" height="22" defaultvalue="&quot;&quot;" tooltip="Descriptive Location Information" group="true" tabstop="true" border="true" required="true" sip="true" field="ALTLOCINFO" multiline="true" hscroll="true"/>
					<LABEL name="Label98876" x="3" y="31" width="47" height="9" caption="Location Info" tooltip="" group="true" border="false"/>
					<EDIT name="txtDateMod" x="9" y="30" width="1" defaultvalue="Now" tooltip="" group="true" tabstop="false" border="false" readonly="true" required="true" sip="false" field="DateMod"/>
					<CHECKBOX name="" x="13" y="29" width="1" defaultvalue="True" caption="CHECKBOX" tooltip="" group="true" tabstop="false" border="false" readonly="true" field="NEEDAUDIT" alignment="left"/>
					<EDIT name="txtInitials" x="3" y="71" width="30" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field="DATAREC"/>
					<LABEL name="lblPDOP" x="107" y="65" width="30" height="9" caption="PDOP label" tooltip="" group="true" border="false"/>
					<EDIT name="txtPDOP" x="73" y="65" width="30" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field="RECPDOP"/>
					<EDIT name="txtSatellites" x="73" y="80" width="30" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field="RECSATS"/>
					<LABEL name="lblSatellites" x="107" y="80" width="20" height="12" caption="SATS" tooltip="" group="true" border="false"/>
					<EDIT name="txtQuality" x="73" y="95" width="30" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field="RECQUAL"/>
					<LABEL name="lblQuality" x="107" y="95" width="30" height="12" caption="QUAL" tooltip="" group="true" border="false"/>
					<EDIT name="txtDifferentialAge" x="73" y="111" width="30" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field="RECDAGE"/>
					<LABEL name="lblDataAge" x="107" y="111" width="30" height="9" caption="DAGE" tooltip="" group="true" border="false"/>
					<EDIT name="txtAccuracy" x="37" y="71" width="30" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field=""/>
				</PAGE>
				<PAGE name="pgDesc" caption="Description" onvalidate="ThisEvent.Result = True">
					<CHECKBOX name="chkActive" x="4" y="1" width="35" height="12" defaultvalue="1" caption="Active" tooltip="" group="true" tabstop="true" border="false" field="ACTIVE" alignment="left"/>
					<LABEL name="Label11" x="39" y="3" width="37" height="10" caption="Recorded" tooltip="" group="true" border="false" alignment="right"/>
					<EDIT name="txtDateRecord" x="80" y="2" width="45" height="12" defaultvalue="date()" tooltip="" group="true" tabstop="true" border="true" sip="true" field="DATERECORD"/>
					<LABEL name="label12" x="3" y="84" width="41" height="10" caption="Disturbance" tooltip="" group="true" border="false"/>
					<COMBOBOX name="cboDisturbance" x="3" y="94" width="89" height="13" defaultvalue="&quot;&quot;" listtable="Disturbances.dbf" listvaluefield="DISTCODE" listtextfield="DISTURB" tooltip="" group="true" tabstop="true" border="false" sip="false" sort="false" field="MAINDISTRB"/>
					<LABEL name="Label13" x="3" y="17" width="43" height="8" caption="Vegetation" tooltip="" group="true" border="false"/>
					<LABEL name="Label15" x="3" y="44" width="23" height="9" caption="Goal" tooltip="" group="true" border="false"/>
					<EDIT name="txtGoal" x="3" y="52" width="121" height="25" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="true" border="true" sip="true" field="GOAL" multiline="true" hscroll="true"/>
					<COMBOBOX name="cboVeg" x="3" y="26" width="123" height="13" defaultvalue="" listtable="Vegetation.dbf" listvaluefield="VEGETATION" listtextfield="VEGETATION" tooltip="" group="true" tabstop="true" border="false" sip="false" field="VEGETATION"/>
				</PAGE>
				<PAGE name="pgLocation" caption="Areas" onvalidate="Call ValidateAreas">
					<COMBOBOX name="cboArea1" x="14" y="1" width="112" height="13" defaultvalue="" listtable="Areas.dbf" listvaluefield="AREAKEY" listtextfield="AREANAME" tooltip="" group="true" tabstop="true" border="false" required="true" sip="false" field="AREA1"/>
					<COMBOBOX name="cboArea2" x="13" y="18" width="113" height="13" defaultvalue="" listtable="Areas.dbf" listvaluefield="AREAKEY" listtextfield="AREANAME" tooltip="" group="true" tabstop="true" border="false" sip="false" field="AREA2"/>
					<COMBOBOX name="cboArea3" x="13" y="37" width="113" height="13" defaultvalue="" listtable="Areas.dbf" listvaluefield="AREAKEY" listtextfield="AREANAME" tooltip="" group="true" tabstop="true" border="false" sip="false" field="AREA3"/>
					<COMBOBOX name="cboArea4" x="13" y="55" width="113" height="13" defaultvalue="" listtable="Areas.dbf" listvaluefield="AREAKEY" listtextfield="AREANAME" tooltip="" group="true" tabstop="true" border="false" sip="false" field="AREA4"/>
					<LABEL name="Label21" y="3" width="10" height="9" caption="1:" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label22" y="18" width="10" height="9" caption="2:" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label23" y="37" width="10" height="9" caption="3:" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label21a" y="55" width="10" height="9" caption="4:" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="lblTownship" x="3" y="74" width="37" height="9" caption="Township" tooltip="" group="true" border="false"/>
					<EDIT name="editTownship" x="40" y="71" width="20" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field=""/>
					<LABEL name="lblRange" x="63" y="74" width="27" height="9" caption="Range" tooltip="" group="true" border="false" alignment="center"/>
					<EDIT name="editRange" x="93" y="71" width="20" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field=""/>
					<LABEL name="lblSection" x="3" y="89" width="37" height="9" caption="Section" tooltip="" group="true" border="false"/>
					<EDIT name="editSection" x="40" y="86" width="20" height="12" defaultvalue="" tooltip="" tabstop="true" border="true" sip="false" field=""/>
				</PAGE>
			</EDITFORM>
		</FORMS>
		<SYSTEMOBJECTS/>
	</LAYER>
	<SCRIPT src="Occurrences.VBS"/>
</ArcPad>

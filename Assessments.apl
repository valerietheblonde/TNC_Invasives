<ArcPad>
	<LAYER name="Assessments" onclose="Call TurnOffApplet">
		<SYMBOLOGY>
			<SIMPLELABELRENDERER visible="false" field="WEEDNAME">
				<TEXTSYMBOL font="Arial" charset="0" fontsize="8.25" fontstyle="regular" fontcolor="0,0,0" angle="0"/>
			</SIMPLELABELRENDERER>
			<SIMPLERENDERER>
				<SIMPLEPOLYGONSYMBOL filltype="solid" filltransparency="0.0" fillcolor="255,255,115" boundarytype="solid" boundarywidth="3.0" boundarycolor="255,170,50" boundarytransparency="1"/>
			</SIMPLERENDERER>
		</SYMBOLOGY>
		<FORMS>
			<EDITFORM name="EDITFORM" caption="Assessment Polygon" width="130" height="130" symbologypagevisible="false" onload="Dim objControls

Set objControls = ThisEvent.Object.Pages(&quot;pgLoc&quot;).Controls
If objControls(&quot;txtVisitID&quot;).Value = &quot;&quot; Then
   objControls(&quot;txtVisitID&quot;).Value = getWeedDbKey
End If
objControls(&quot;txtDateMod&quot;).Value = Now
If Not objControls(&quot;txtDate_&quot;).Value = &quot;&quot; Then
   Dim dDate
   dDate = cDate(objControls(&quot;txtDate_&quot;).Value)
   objControls(&quot;dtpDate&quot;).Value = dDate
End If
Set objControls = Nothing
If Not Map.PointerMode = &quot;modeidentify&quot; Then
	Call PopulateWO
End If">
				<PAGE name="pgLoc" caption="Location" onkillactive="Call ValidateWeed()">
					<DATETIME name="dtpDate" x="61" y="85" width="65" height="13" defaultvalue="" onchange="Dim page
Set page = ThisEvent.Object.Parent
page.Controls.Item(&quot;txtDate_&quot;).Value = FormatDateTime(ThisEvent.Object.Value, vbShortDate)
Set page = Nothing" tooltip="" group="true" tabstop="true" border="true" required="true" sip="false" field="" allownulls="false"/>
					<COMBOBOX name="cboOccurrence" x="1" y="10" width="127" height="13" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" onselchange="Call FillWOFields(ThisEvent.Object.Value)" tooltip="" group="true" tabstop="true" border="false" sip="false" field="WOKEY"/>
					<EDIT name="txtWeedName" x="43" y="25" width="83" height="12" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="true" readonly="true" sip="false" field="WEEDNAME"/>
					<LABEL name="label1" x="5" y="1" width="83" height="8" caption="Choose Occurrence:" tooltip="" group="true" border="false"/>
					<LABEL name="label2" x="20" y="26" width="20" height="10" caption="Weed" tooltip="" group="true" border="false" alignment="right"/>
					<EDIT name="txtCrew" x="43" y="36" width="83" height="12" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="true" sip="false" field="CREW"/>
					<LABEL name="label6" x="22" y="37" width="20" height="10" caption="Crew" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="LAbel7" x="59" y="76" width="18" height="8" caption="Date" tooltip="" group="true" border="false"/>
					<DATETIME name="dtpDateMod" x="5" y="5" width="1" height="1" defaultvalue="Now" tooltip="" group="true" tabstop="false" border="false" required="true" sip="false" field="DATEMOD" allownulls="false"/>
					<EDIT name="txtNotes" x="3" y="50" width="123" height="24" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="true" border="true" sip="true" field="NOTES" multiline="true" hscroll="true"/>
					<LABEL name="label8" x="2" y="40" width="21" height="9" caption="Notes" tooltip="" group="true" border="false"/>
					<EDIT name="txtVisitID" x="1" y="1" width="1" height="1" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="false" border="true" readonly="true" required="true" sip="false" field="VISITID"/>
					<EDIT name="txtDateMod" x="13" y="41" width="1" height="1" defaultvalue="Now" tooltip="" group="true" tabstop="false" border="false" required="true" sip="false" field="DATEMOD"/>
					<EDIT name="txtDate_" x="14" y="60" width="1" height="1" defaultvalue="FormatDateTime(Now, vbShortDate)" tooltip="" group="true" tabstop="false" border="false" required="true" sip="false" field="DATE_" multiline="true" hscroll="true"/>
					<COMBOBOX name="txtDataRec" x="3" y="85" width="41" height="13" defaultvalue="&quot;&quot;" listtable="WeedassessmentsP.dbf" listvaluefield="DATAREC" listtextfield="DATAREC" tooltip="" group="true" tabstop="true" border="false" limittolist="false" field="DATAREC"/>
					<LABEL name="lblYou" x="2" y="76" width="19" height="8" caption="You" tooltip="" group="true" border="false"/>
				</PAGE>
				<PAGE name="pgSize" caption="Size">
					<COMBOBOX name="cboAccuracy" x="57" y="2" width="47" height="13" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" tooltip="" group="true" tabstop="true" border="false" sip="false" sort="false" field="ACCURACY">
						<LISTITEM value="none" text="none"/>
						<LISTITEM value="GPS1" text="GPS1"/>
						<LISTITEM value="GPS2" text="GPS2"/>
						<LISTITEM value="GPS3" text="GPS3"/>
						<LISTITEM value="MAN1" text="MAN1"/>
						<LISTITEM value="MAN2" text="MAN2"/>
						<LISTITEM value="MAN3" text="MAN3"/>
					</COMBOBOX>
					<LABEL name="lblAcc" x="19" y="4" width="35" height="10" caption="Accuracy" tooltip="" group="true" border="false" alignment="right"/>
					<CHECKBOX name="chkCalcFromShape" x="8" y="79" width="103" height="12" defaultvalue="0" onclick="Call CalcSize" caption="Calculate size from shape" tooltip="" group="true" tabstop="true" border="false" field="" alignment="left" allownulls="false"/>
					<EDIT name="txtLength" x="57" y="19" width="65" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;Length&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgSize&quot;).Controls(&quot;txtLength&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="false" field="NUMLENGTH"/>
					<EDIT name="txtWidth" x="57" y="33" width="66" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;Width&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgSize&quot;).Controls(&quot;txtWidth&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="false" field="NUMWIDTH"/>
					<COMBOBOX name="cboUOM" x="57" y="47" width="45" height="13" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" tooltip="" group="true" tabstop="true" border="false" sip="false" limittolist="false" sort="false" field="LENWIDUOM">
						<LISTITEM value="ft" text="ft"/>
						<LISTITEM value="m" text="m"/>
						<LISTITEM value="km" text="km"/>
						<LISTITEM value="mi" text="mi"/>
					</COMBOBOX>
					<EDIT name="txtAcres" x="57" y="66" width="58" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;Area&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgSize&quot;).Controls(&quot;txtAcres&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="false" field="GROSSACRES"/>
					<LABEL name="Label31" x="21" y="20" width="33" height="10" caption="Length" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label32" x="27" y="34" width="29" height="10" caption="Width" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label33" x="31" y="49" width="25" height="10" caption="UOM" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label34" x="13" y="67" width="42" height="10" caption="Area (in ac)" tooltip="" group="true" border="false" alignment="right"/>
				</PAGE>
				<PAGE name="pgCovDen" caption="Cover/Dens">
					<LABEL name="Label41" x="5" y="4" width="50" height="10" caption="'Cover' Info" tooltip="" group="true" border="false" fontsize="9" fontstyle="bolditalic"/>
					<LABEL name="Label42" x="7" y="16" width="41" height="10" caption="Percent" tooltip="" group="true" border="false" alignment="right"/>
					<EDIT name="txtPctCover" x="52" y="14" width="67" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;Cover Percent&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgCovDen&quot;).Controls(&quot;txtPctCover&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="true" field="COVERPERC"/>
					<LABEL name="Label43" x="29" y="31" width="20" height="10" caption="Class" tooltip="" group="true" border="false" alignment="right"/>
					<COMBOBOX name="cboCovClass" x="52" y="30" width="68" height="13" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" onselchange="Call CheckCover()" tooltip="" group="true" tabstop="true" border="true" sip="false" sort="false" field="COVERCLS">
						<LISTITEM value="&lt; 1%" text="&lt; 1%"/>
						<LISTITEM value="1 - 10%" text="1 - 10%"/>
						<LISTITEM value="11 - 25%" text="11 - 25%"/>
						<LISTITEM value="26 - 50%" text="26 - 50%"/>
						<LISTITEM value="51 - 100%" text="51 - 100%"/>
					</COMBOBOX>
					<LABEL name="Label44" x="4" y="50" width="50" height="10" caption="'Density' Info" tooltip="" group="true" border="false" fontsize="9" fontstyle="bolditalic"/>
					<LABEL name="Label45" x="11" y="63" width="37" height="10" caption="Density" tooltip="" group="true" border="false" alignment="right"/>
					<EDIT name="txtDensity" x="52" y="62" width="60" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;Density&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgCovDen&quot;).Controls(&quot;txtDensity&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="true" field="DENSUNIT"/>
					<LABEL name="Label46" x="9" y="78" width="41" height="10" caption="Unit Area" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label47" x="2" y="94" width="27" height="10" caption="Count" tooltip="" group="true" border="false" alignment="right"/>
					<COMBOBOX name="cboDensCount" x="31" y="92" width="85" height="13" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" tooltip="" group="true" tabstop="true" border="false" sip="false" sort="false" field="DENSCNTUNI">
						<LISTITEM value="All plants, w/o seedlings" text="All plants, w/o seedlings"/>
						<LISTITEM value="All plants w/ seedlings" text="All plants w/ seedlings"/>
						<LISTITEM value="Only flowering plants" text="Only flowering plants"/>
						<LISTITEM value="All stems w/o seedlings" text="All stems w/o seedlings"/>
						<LISTITEM value="All stems w/ seedlings" text="All stems w/ seedlings"/>
					</COMBOBOX>
					<COMBOBOX name="cboDensUnitArea" x="52" y="76" width="63" height="13" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" tooltip="" group="true" tabstop="true" border="true" sip="false" sort="false" field="DENSITYUOM">
						<LISTITEM value="sq. ft" text="sq. ft"/>
						<LISTITEM value="sq. meters" text="sq. meters"/>
						<LISTITEM value="Acres" text="Acres"/>
						<LISTITEM value="Hectares" text="Hectares"/>
						<LISTITEM value="Infested Area" text="Infested Area"/>
						<LISTITEM value="# plants" text="# plants"/>
					</COMBOBOX>
					<BUTTON onclick="Msgbox &quot;If Cover % is ZERO OR BLANK, choosing a Cover Class will plug an average value into Cover %.&quot;" name="btnCoverHelp" x="3" y="28" width="25" height="14" caption="Help" tooltip="" group="true" tabstop="true" border="false" alignment="center"/>
				</PAGE>
				<PAGE name="pgStat" caption="Stats">
					<EDIT name="txtComments" x="3" y="10" width="122" height="24" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="true" border="true" sip="true" field="COMMENTS" multiline="true" hscroll="true"/>
					<COMBOBOX name="cboPheno" x="45" y="36" width="80" height="100" defaultvalue="&quot;&quot;" listtable="Phenology.dbf" listvaluefield="PHENOCODE" listtextfield="PHENO" tooltip="" group="true" tabstop="true" border="true" sip="false" sort="false" field="PHENO"/>
					<LABEL name="label11" x="2" y="37" width="41" height="10" caption="Phenology" tooltip="" group="true" border="false" alignment="right"/>
					<COMBOBOX name="cboSTTrend" x="45" y="49" width="80" height="100" defaultvalue="&quot;&quot;" listtable="Trends.dbf" listvaluefield="TREND" listtextfield="TREND" tooltip="" group="true" tabstop="true" border="true" sip="false" sort="false" field="STTREND"/>
					<LABEL name="label12" x="4" y="51" width="40" height="10" caption="STTrend" tooltip="" group="true" border="false" alignment="right"/>
					<COMBOBOX name="cboLTTrend" x="46" y="64" width="80" height="100" defaultvalue="&quot;&quot;" listtable="Trends.dbf" listvaluefield="TREND" listtextfield="TREND" tooltip="" group="true" tabstop="true" border="true" sip="false" sort="false" field="LTTREND"/>
					<LABEL name="label13" x="11" y="66" width="34" height="12" caption="LTTrend" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="Label14" x="3" y="79" width="42" height="10" caption="Distribution" tooltip="" group="true" border="false" alignment="right"/>
					<COMBOBOX name="cboDistribution" x="46" y="79" width="80" height="100" defaultvalue="&quot;&quot;" listtable="" listvaluefield="" listtextfield="" tooltip="" group="true" tabstop="true" border="true" sip="false" sort="false" field="DISTRIB">
						<LISTITEM value="I" text="Isolated"/>
						<LISTITEM value="L" text="Linear"/>
						<LISTITEM value="M" text="Monoculture"/>
						<LISTITEM value="S" text="Satellite"/>
						<LISTITEM value="U" text="Uniform"/>
						<LISTITEM value="O" text="Other"/>
					</COMBOBOX>
					<LABEL name="Label16" x="2" y="1" width="35" height="8" caption="Comments" tooltip="" group="true" border="false" alignment="right"/>
				</PAGE>
				<PAGE name="pgTime" caption="Time">
					<LABEL name="Label21" width="104" height="10" caption="Record EITHER start/end" tooltip="" group="true" border="false" fontsize="9" fontstyle="bolditalic"/>
					<LABEL name="label22" x="1" y="9" width="103" height="10" caption="times OR total time" tooltip="" group="true" border="false" fontsize="9" fontstyle="bolditalic"/>
					<EDIT name="txtTimeStart" x="10" y="29" width="40" height="12" defaultvalue="&quot;&quot;" tooltip="" group="true" tabstop="true" border="true" sip="false" field="TIMESTART"/>
					<EDIT name="txtTimeEnd" x="60" y="29" width="40" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateTime()" tooltip="" group="true" tabstop="true" border="true" sip="false" field="TIMEEND"/>
					<LABEL name="label23" x="9" y="19" width="41" height="10" caption="Start Time" tooltip="" group="true" border="false"/>
					<LABEL name="label24" x="59" y="19" width="50" height="10" caption="End Time" tooltip="" group="true" border="false"/>
					<EDIT name="txtTotHours" x="49" y="44" width="33" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;Total Time&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgTime&quot;).Controls(&quot;txtTotHours&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="false" field="TOTHOURS"/>
					<LABEL name="Label15" x="11" y="46" width="36" height="10" caption="Total Time" tooltip="" group="true" border="false" alignment="right"/>
					<EDIT name="txtNumStaff" x="32" y="70" width="30" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;# Staff&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgTime&quot;).Controls(&quot;txtNumStaff&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="false" field="NUMSTAFF"/>
					<LABEL name="Label26" x="32" y="60" width="30" height="8" caption="Staff" tooltip="" group="true" border="false"/>
					<LABEL name="Label28" x="77" y="62" width="39" height="8" caption="Volunteer" tooltip="" group="true" border="false"/>
					<EDIT name="txtNumVols" x="77" y="70" width="30" height="12" defaultvalue="&quot;&quot;" onkillfocus="Call ValidateNumeric(&quot;# Volunteers&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgTime&quot;).Controls(&quot;txtNumVols&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" sip="false" field="NUMVOLS"/>
					<BUTTON onclick="msgbox(&quot;Key times in 24-hour format, like 0830.  Key hours like 4.5&quot;)" name="btnHelp" x="104" y="2" width="22" height="14" caption="Help" tooltip="" group="true" tabstop="true" border="false" alignment="center"/>
					<LABEL name="lblDash" x="52" y="31" width="5" height="8" caption="-" tooltip="" group="true" border="false" fontsize="9" fontstyle="bolditalic" alignment="center"/>
					<LABEL name="lblStaffHrs" x="12" y="71" width="17" height="10" caption="#" tooltip="" group="true" border="false" alignment="right"/>
					<LABEL name="lblVolHrs" x="5" y="85" width="26" height="10" caption="Hours" tooltip="" group="true" border="false" alignment="right"/>
					<EDIT name="StaffHrs" x="32" y="83" width="29" height="12" defaultvalue="" onkillfocus="Call ValidateNumeric(&quot;Staff Hours&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgTime&quot;).Controls(&quot;StaffHrs&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" field="STAFFHOURS"/>
					<EDIT name="VolHrs" x="77" y="83" width="29" height="12" defaultvalue="" onkillfocus="Call ValidateNumeric(&quot;Volunteer Hours&quot;,Map.Layers(&quot;Assessment Polygons&quot;).Forms(&quot;EDITFORM&quot;).Pages(&quot;pgTime&quot;).Controls(&quot;VolHrs&quot;).Value)" tooltip="" group="true" tabstop="true" border="true" field="VOLHOURS"/>
				</PAGE>
			</EDITFORM>
		</FORMS>
		<SYSTEMOBJECTS/>
	</LAYER>
	<SCRIPT src="Assessments.vbs" language="VBScript"/>
</ArcPad>

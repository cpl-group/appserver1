<table border=0 cellpadding="0" cellspacing="0">
          <tr>
         <% '12/20/2007 N.Ambo added criteria to only show edit button for users in group 'Job Status Admins' if the job status is 'closed'
         if (cStatus = "Closed") then 
				if allowgroups("Job Status Admins") then %>
					  <td><input id="editjob" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="edit_job('<%=jid%>');" value="Edit Job">
					&nbsp;</td>		 
				<% end if %>
            <% elseif (cStatus<>"Closed" or allowgroups("Genergy_Corp,Joblog_Admin")) then %>
            <td><input id="editjob" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="edit_job('<%=jid%>');" value="Edit Job">
              &nbsp;</td>
            <%end if%>
            <td><input name="opentt" type="button" class="standard" id="opentt" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="window.open('/genergy2_intranet/itservices/ttracker/ticket.asp?bldg=&mode=new&jobid=<%=jid%>&child=1&ticketfortype=joblog&ticketfor=<%=jid%>','MeterTroubleTicket','width=680,height=325')" value="Open a Trouble Ticket"></td>
            <% if lcase(cStatus)="in progress" then %>
            <td><input id="invoice" name="invoice" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="viewinvoice('<%=jid%>');" value="Invoice Job">
              &nbsp;</td>
            <td><input id="new_po" name="new_po" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="newPO();" value="New Requisition Form">
              &nbsp;</td>
            <td><input id="addtask" name="addtask" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="newTask('<%=jid%>');" value="Add Task">
              &nbsp;</td>
			<% 
			Dim showChangeOrderButton
			if showChangeOrderButton then %>
<td><input id="changeorder" name="changeorder" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="newchange();" value="New Change Order">
              &nbsp;</td>			<%end if %>
            <% end if %>
          </tr>
        </table>
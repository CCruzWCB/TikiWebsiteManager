<%@ Page Language="VB" MasterPageFile="~/MasterPages/NonSecure.master" AutoEventWireup="false" CodeFile="CreateUser.aspx.vb" Inherits="NonSecure_CreateUser" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
    <asp:CreateUserWizard ID="CreateUserWizard1" runat="server" MembershipProvider="AspNetSqlMembershipProvider">
        <WizardSteps>
            <asp:CreateUserWizardStep runat="server">
            </asp:CreateUserWizardStep>
            <asp:CompleteWizardStep runat="server">
            </asp:CompleteWizardStep>
        </WizardSteps>
        <FinishNavigationTemplate>
            <asp:Button ID="FinishPreviousButton" runat="server" CausesValidation="False" CommandName="MovePrevious"
                Text="Previous" />
            <asp:Button ID="FinishButton" runat="server" CommandName="MoveComplete" Text="Finish" />
        </FinishNavigationTemplate>
    </asp:CreateUserWizard>

</asp:Content>


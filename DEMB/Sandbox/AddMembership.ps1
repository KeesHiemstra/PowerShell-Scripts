$Engineer = @('david.konvicny', 'petr.beno', 'jiri.ovcacek')

#Add HP-Deskside-EMEA membership

foreach ( $Item in $Engineer )
{
    Add-ADGroupMember -Identity 'HP-Deskside-EMEA' -Members (Get-ADUser $Item) -Confirm:$false
}

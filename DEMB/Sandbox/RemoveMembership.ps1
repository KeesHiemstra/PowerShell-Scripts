$Engineer = @('david.konvicny', 'petr.beno', 'jiri.ovcacek')

#Remove HP-Deskside-EMEA membership

foreach ( $Item in $Engineer )
{
    Remove-ADGroupMember -Identity 'HP-Deskside-EMEA' -Members (Get-ADUser $Item) -Confirm:$false
}

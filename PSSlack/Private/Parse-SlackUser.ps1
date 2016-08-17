﻿# Parse users
Function Parse-SlackUser {
    [cmdletbinding()]
    param( $InputObject )

    foreach($User in $InputObject)
    {

        [pscustomobject]@{
            PSTypeName = 'PSSlack.User'
            ID = $User.id
            Name = $User.name
            RealName = $User.Profile.Real_Name
            FirstName = $User.Profile.First_Name
            Last_Name = $User.Profile.Last_Name
            Email = $User.Profile.email
            Phone = $User.Profile.Phone
            Skype = $User.Profile.Skype
            IsBot = $User.Is_Bot
            IsAdmin = $User.Is_Admin
            IsOwner = $User.Is_Owner
            IsPrimaryOwner = $User.Is_Primary_Owner
            IsRestricted = $User.Is_Restricted
            IsUltraRestricted = $User.Is_Ultra_Restricted
            Status = $User.Status
            TimeZoneLabel = $User.tz_label
            TimeZone = $User.tz
            Presence = $User.Presence
            Deleted = $User.Deleted
            Raw = $User
        }
    }
}
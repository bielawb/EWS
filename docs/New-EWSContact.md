---
external help file: EWS-help.xml
online version: 
schema: 2.0.0
---

# New-EWSContact

## SYNOPSIS
Creates new contacts using provided contact details.

## SYNTAX

```
New-EWSContact [[-GivenName] <String>] [[-Surname] <String>] [[-MiddleName] <String>] [[-Company] <String>]
 [[-Phone] <Hashtable>] [[-Email] <Hashtable[]>] [[-PhysicalAddress] <Hashtable>]
 [[-Service] <ExchangeService>]
```

## DESCRIPTION
Function that helps creating new contacts.
It accepts single hash table for phone and physical address with details of the contact.
Keys of these hash tables need to match with the types defined for a given contact, e.g.:
Phone: HomePhone, MobilePhone
PhysicalAddress: Home, Business

Emails, on the other hand, accepts array of hash tables, each with two keys: DisplayName and email.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> New-EWSContact -GivenName John -Surname Doe
```

Create new contact for 'John Doe' without any details populated.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> $mails = @{ DisplayName = 'John Doe'; email = 'johndoe@live.com' }
PS C:\> $phones = @{ MobilePhone = '+48 601 602 603'; HomePhone = '+48 22 123 4567' }
PS C:\> New-EWSContact -GivenName John -Surname Doe -Email $mails -Phone $phones
```

Create new contact for 'John Doe' with e-mail and two phone types specified.

## PARAMETERS

### -Company
Name of the contact's company.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Email
Collection of hash tables with two keys: Email and DisplayName.

```yaml
Type: Hashtable[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GivenName
First name of the contact.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MiddleName
Middle name of the contact.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Phone
Hash table with phone details of the contact, each key has to match to phone type.

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PhysicalAddress
Hash table with physical address details of the contact, each key has to match to physical address type.
Each value is another hash table with keys matching to properties of the address.

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases: 

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Service
Service used to create new contact.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases: 

Required: False
Position: 7
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Surname
Last name of the contact.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

### Microsoft.Exchange.WebServices.Data.ExchangeService


## OUTPUTS

### Microsoft.Exchange.WebServices.Data.Contact


## NOTES

## RELATED LINKS


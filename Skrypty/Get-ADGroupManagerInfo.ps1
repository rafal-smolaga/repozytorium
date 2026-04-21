# Nazwa grupy do pobrania (do modyfikacji)
$groupName = "Nazwa grupy"

# Pobranie grupy z wszystkimi potrzebnymi właściwościami
$group = Get-ADGroup -Filter {Name -eq $groupName} -Properties Description, Info, managedBy

# Pobranie danych zarządcy jeśli istnieje
if ($group.managedBy) {
    $manager = Get-ADUser -Identity $group.managedBy -Properties Name
    $managerName = $manager.Name
} else {
    $managerName = 'N/A'
}

# Wyświetlenie wyników
$group | Select-Object @{n='Group Name';e={$group.Name}},
                        @{n='Managed By Name';e={$managerName}},
                        @{n='Description';e={$group.Description}},
                        @{n='Notes';e={$group.Info}}
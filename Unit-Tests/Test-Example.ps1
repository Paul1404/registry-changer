Describe "Registry-Changer.ps1" {
    BeforeAll {
        # Set up any necessary test environment
    }

    AfterAll {
        # Clean up any test artifacts
        # For example, you might want to delete the test registry key or restore the original registry state
    }

    Context "Get-NewParameters" {
        It "Should return valid parameters" {
            # Mock the user input for Read-Host
            $mockInput = @{
                'Enter the registry path' = 'HKLM:\SOFTWARE\TestKey'
                'Enter the value name' = 'TestValue'
                'Enter the value data' = 'Test Data'
            }
            Mock -CommandName Read-Host -MockWith { param($prompt) $mockInput[$prompt] }

            # Call the function and capture the output
            $result = Get-NewParameters

            # Assert the expected result
            $expectedResult = [PSCustomObject]@{
                RegPath = 'HKLM:\SOFTWARE\TestKey'
                ValueName = 'TestValue'
                ValueData = 'Test Data'
            }
            $result | Should Be $expectedResult
        }
    }

    Context "Update-UserProfiles" {
        It "Should update registry settings for user profiles" {
            # Set up any necessary test environment

            # Call the function and capture the output
            $result = Update-UserProfiles -parameters $parameters

            # Assert the expected result

            # Validate that the registry settings were updated correctly for user profiles
            Write-Output $result
            # Clean up any test artifacts
        }
    }

    # Placeholder for Testing Functions
    
}

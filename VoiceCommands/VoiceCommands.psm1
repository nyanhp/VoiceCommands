function Out-Voice
{
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [System.String[]]
        $Text,

        [ValidateSet('Male', 'Female')]
        [System.String]
        $Gender = 'Female',

        [cultureinfo]
        $VoiceCulture = 'en-us',

        [int16]
        [ValidateRange(-10,10)]
        $Rate = 0
    )

    begin
    {
        $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $synth.SelectVoiceByHints($Gender, 30, $null, $VoiceCulture)
        $synth.SetOutputToDefaultAudioDevice()
        $synth.Rate = $Rate
    }
    
    process
    {  
        foreach ( $t in $text)
        {
            $synth.Speak($t)
        }
    }

    end
    {
        $synth.Dispose()
    }
}

New-Alias -Name ov -Value Out-Voice -Description "Alias for voice cmdlet Out-Voice" -Force

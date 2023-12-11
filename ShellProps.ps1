#
# Use shell extended properties to map photo GPS coordinates to a location
# Photos can then be renamed based on the location they were taken from

#
# Sample output
#
# ImageName               DateTaken              Latitude Longitude Location                      
# ---------               ---------              -------- --------- --------                      
# IMG_20141128_171525.jpg 11/28/2014 4:15:25 PM   49.2694 -123.1237 Canada/Vancouver              
# IMG_20141026_160520.jpg 10/26/2014 3:05:20 PM   48.4181 -123.3718 Canada/Victoria               
# IMG_20150724_114616.jpg 7/24/2015 9:46:16 AM    48.7111   -1.8446 France/Cancale                
# IMG_20150804_164235.jpg 8/4/2015 2:42:35 PM     48.6056   -2.1549 France/Lancieux               
# IMG_20150806_112731.jpg 8/6/2015 9:27:31 AM     48.6014   -2.0707 France/Pleurtuit              
# IMG_20150804_131257.jpg 8/4/2015 11:12:57 AM    48.6681   -2.2848 France/Plévenon               
# IMG_20150804_124959.jpg 8/4/2015 10:49:59 AM    48.6372   -2.2552 France/Saint-Cast-le-Guildo   
# IMG_20150804_102431.jpg 8/4/2015 8:24:31 AM     48.6539   -2.0039 France/Saint-Malo             
# IMG_20141220_120444.jpg 12/20/2014 11:04:44 AM  20.3659  -87.3326 México/city not found         
# IMG_20141222_121150.jpg 12/22/2014 11:11:50 AM  20.3795  -87.3295 México/Puerto Aventuras       
# IMG_20141223_110744.jpg 12/23/2014 10:07:44 AM  20.6981  -88.5701 México/Tinum                  
# IMG_20150304_173941.jpg 3/4/2015 4:39:41 PM     47.6197 -122.1697 United States/Bellevue        
# IMG_20150823_101619.jpg 8/23/2015 8:16:19 AM    47.5755 -120.6136 United States/Chelan County   
# IMG_20150411_133816.jpg 4/11/2015 11:38:16 AM    32.688 -117.1789 United States/Coronado        
# IMG_20150410_130306.jpg 4/10/2015 11:03:06 AM   33.0388  -117.273 United States/Encinitas       
# IMG_20141026_163616.jpg 10/26/2014 3:36:16 PM   47.9274 -122.5321 United States/Hansville       
# IMG_20150523_122034.jpg 5/23/2015 10:20:34 AM   47.5525 -122.0499 United States/Issaquah        
# IMG_20141026_163049.jpg 10/26/2014 3:30:49 PM   48.1608 -122.7608 United States/Jefferson County
# IMG_20151003_151151.jpg 10/3/2015 1:11:51 PM    47.6422 -122.2403 United States/King County     
# IMG_20141026_164209.jpg 10/26/2014 3:42:09 PM   47.7487 -122.4452 United States/Kingston        
# IMG_20150519_192753.jpg 5/19/2015 5:27:53 PM    47.6689 -122.2022 United States/Kirkland        
# IMG_20150410_144426.jpg 4/10/2015 12:44:26 PM   32.8447 -117.2781 United States/La Jolla        
# IMG_20150816_163713.jpg 8/16/2015 2:37:13 PM    47.3716 -122.0387 United States/Maple Valley    
# IMG_20140903_081854.jpg 9/3/2014 6:18:54 AM     47.5821 -122.2064 United States/Mercer Island   
# IMG_20150621_114252.jpg 6/21/2015 9:42:52 AM    45.6032 -122.6721 United States/Portland        
# IMG_20141124_091356.jpg 11/24/2014 8:13:56 AM   47.6491 -122.1423 United States/Redmond         
# IMG_20150412_122005.jpg 4/12/2015 10:20:05 AM   32.7079 -117.1745 United States/San Diego       
# IMG_20150406_184049.jpg 4/6/2015 4:40:49 PM     47.5858 -122.3342 United States/Seattle

      
Param(
    [string]$Folder = "C:\Users\Phili\OneDrive\Pictures\Camera Roll"
)

# Get Shell object
$sh = New-Object -COMObject Shell.Application
$shfolder = $sh.Namespace($Folder)

foreach($f in (Get-ChildItem -Path $Folder | Where-Object Name -match '\.jpg$|\.jpeg$')) {

    # Get a shell file object
    $shf=$shfolder.ParseName($f)

    # Get the properties (see https://learn.microsoft.com/en-us/windows/win32/properties/props)

    $lat=$shf.ExtendedProperty("System.GPS.Latitude")
    $declat=$lat[0] + $lat[1]/60.0 + $lon[2]/3600.0
    if($shf.ExtendedProperty("System.GPS.LatitudeRef") -eq 'S') {
        $declat=-$declat
    }

    $lon=$shf.ExtendedProperty("System.GPS.Longitude")
    $declon=$lon[0] + $lon[1]/60.0 + $lon[2]/3600.0
    if($shf.ExtendedProperty("System.GPS.LongitudeRef") -eq 'W') {
        $declon=-$declon
    }

    $datetaken=$shf.ExtendedProperty("System.Photo.DateTaken")

    if($null -ne $declat -and $null -ne $declon) {
        # Call into free/no-API key GPS mapping service
        # Can alternatively get an API key from Bing or Google account
        $uri="https://geocode.maps.co/reverse?lat=$declat&lon=$declon"
        $loc=Invoke-RestMethod -Uri $uri

        # Some mappings return town, other city
        if($loc.address.PSobject.Properties.name -match "town") {
            $townorcity=$loc.address.town
        } elseif($loc.address.PSobject.Properties.name -match "city") {
            $townorcity=$loc.address.city
        } elseif($loc.address.PSobject.Properties.name -match "county") {
            $townorcity=$loc.address.county
        } else {
            $townorcity="city not found"
        }

        [PsCustomObject]@{
            ImageName = $f.Name
            DateTaken = $datetaken
            Latitude  = [Math]::round($declat, 4)
            Longitude = [Math]::round($declon, 4)
            Location  = $loc.address.country + '/' + $townorcity
        }

        # 2 calls/sec to avoid geocode throttling
        Start-Sleep -Milliseconds 500
    }
}


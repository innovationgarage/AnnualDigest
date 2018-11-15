Replaces text strings inside a presentation

    Usage:
    RenderToVideo.exe input_powerpoint.pptx output_video.mp4
    
By default generates a High bitrate 4K 45 FPS video (hardcoded). It seems the PowerPoint renderer is quite buggy when using non standard parameters. For example, testing different framerates:

| FPS | Results |
|-----|---------|
|30|**Good output**, but jagged animations|
|45|**Almost perfect output**, but not standard framerate|
|55|Transition animations get super garbled and glitched|
|59|Almost everything get super garbled and glitched|
|60|Looks ok, but there is an offset in every object after a transition happens|
|89|Transition animations get super garbled and glitched|
|120|Transition animations get super garbled and glitched|

Requires:
* Microsoft PowerPoint 16.0 Object Library
* Microsoft Office 16.0 Object Library

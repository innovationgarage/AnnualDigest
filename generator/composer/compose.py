from moviepy.editor import *
from moviepy.audio.fx.all import *

clip = VideoFileClip("video45.mp4",audio=False).set_fps(60)
audio = AudioFileClip("Up_on_the_Housetop_Instrumental.mp3")


c1 = ColorClip((clip.w,clip.h),duration=2,color=(0,0,0))
clip2 =CompositeVideoClip([c1,clip.set_start(c1.end-1).crossfadein(1)])
clip2.audio = audio.subclip(0,clip2.end).audio_fadeout(10)
 

print("Writing video.")

clip2.write_videofile("video45_audio4.mp4")
print("Done.")

#result = CompositeAudioClip()
#result.write_videofile("editedvideo.webm",fps=25) # Many options...

from moviepy.editor import *
from moviepy.video.fx.crop import crop

import sys

clip = VideoFileClip(sys.argv[1],audio=False) #.set_fps(60)
audio = AudioFileClip(sys.argv[2])

c1 = ColorClip((clip.w,clip.h),duration=2,color=(0,0,0))
clip2 =CompositeVideoClip([c1,clip.set_start(c1.end-1).crossfadein(4)])
clip2.audio = audio.subclip(0,clip2.end).audio_fadeout(10)
 
print("Writing video.")

clip2.write_videofile(sys.argv[3])
print("Done.")

#result = CompositeAudioClip()
#result.write_videofile("editedvideo.webm",fps=25) # Many options...
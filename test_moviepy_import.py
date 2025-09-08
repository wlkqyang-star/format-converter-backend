import sys
print(sys.path)
try:
    from moviepy.editor import VideoFileClip
    print("moviepy.editor imported successfully")
except ImportError as e:
    print(f"ImportError: {e}")



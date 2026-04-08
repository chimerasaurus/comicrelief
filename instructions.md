# Background
I collect digital comic books because I enjoy them. I think my comic book metadata is all over the place though. Specifically, I suspect my comic books have inconsistent, missing, or wrong metadata embedded in them. I recall that cbz and cbr files have metadata embedded. I am also not sure what metadata format (I recall there are two?) is common right now and how to fix it. Beyond that, the file names are all inconsistent. All of this makes it hard to read them with common software because they show up as different series, or out of order, or have missing dates, and so on.

# What I want
A magical Python script I can point at a folder of comics, that grabs the appropriate comic metadata, and applies it (fixes) the individual comic books. It can infer, I think, from the folder and file names (or existing metadata.) Knowing when it runs into ambigious issues, or errors, would also be useful.

I want the script to confirm EACH time a file is changed, showing me like a 2 box with data about the file at the top and then the left of the two box showing the old data and the right box showing the new data.

# Examples
* /Volumes/library/Fiction/Comics/Star.Trek.Starfleet.Academy.\(1996\)
* /Volumes/library/Fiction/Comics

I keep all of my comics in that last directory, so it's a LOT of comics. 

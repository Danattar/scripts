use strict;
use warnings;
#!/usr/bin/perl
sub initialize{
print "Hi, welcome to the installation program of automatic sort of your films.\n\n";
}
sub getDir{
    print "Please enter the name of the folder which is used to receive your downloads :\n";
    my $dirExists = ('');
    my $dir;

#    while($dirExists == ('')){
    while(!$dirExists){
#        $dir = <>;
        $dir = 'C:\perl\test';
        chomp $dir;
        $dir =~ s/\R//g; #enleve tous  les caractères de retour à la ligne dans la variable.
        if (-d $dir){
            # directory called cgi-bin exists
            $dirExists=('1');
           print "Dir $dir Exists\n";
        }
        elsif (-e $dir) {
            # cgi-bin exists but is not a directory
            print "Directory $dir is not a directory.\n";
        }
        else {
            # nothing called cgi-bin exists
            print "No dir, Please reenter the name of directory.\n\n\n";
        }
    }
    return $dir;
}
sub getFilesFromDir{
    my $directory = $_[0];
    my @sorted_files = ();
    print "Please wait while files are retrieve from \"$directory\"\n";
    #open (DIR, $directory);
    opendir(DIR, $directory) or die "opendir() failed: $!";
    my @dir_files = readdir(DIR) or die "error\n";
    closedir(DIR);
    print "Files retrieved in $directory are :\n\n";
    foreach my $file (@dir_files){
        if($file =~ /\.(torrent|mkv)$/i){
            print "$file\n";
            push @sorted_files, $file;
        }
    }
    print "\n";
    return @sorted_files;
}
initialize();
my $dir = getDir();
my @files = getFilesFromDir($dir);

#todo, use imdb api to populate film DB

package Win32::CaptureIE;

use 5.006;
use strict;
use warnings;

require Exporter;

our @ISA = qw(Exporter);

our %EXPORT_TAGS = ( 'default' => [ qw(
  StartIE
  QuitIE
  Navigate
  Refresh
  GetElement
  GetDoc

  CaptureElement
  CapturePage
  CaptureBrowser

  $IE
  $Doc
  $Body
  $HWND_IE
  $HWND_Browser
  $CaptureBorder
) ] );

$EXPORT_TAGS{all} = [ map {@$_} values %EXPORT_TAGS ];

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } );

our @EXPORT = @{ $EXPORT_TAGS{'default'} };

our $VERSION = '1.00';
our $IE;
our $HWND_IE;
our $HWND_Browser;
our $Doc;
our $Body;

our $CaptureBorder = 1;

# Preloaded methods go here.

use Win32::OLE qw(in valof EVENTS);
use Win32::Screenshot qw(:all);
use POSIX qw(ceil floor);
use strict;

##########################################################################

# HACK: DocumentComplete event is not raised when refreshing page but
# only DownloadComplete, so if refreshing page we need to wait for
# DownloadComplete but not if we are navigating to a page
our $refreshing_page = 0;

sub StartIE () {
  my %arg = @_;

  # Open a new browser window and save its' window handle
  $IE = Win32::OLE->new("InternetExplorer.Application");
  Win32::OLE->WithEvents($IE,\&EventHandler,"DWebBrowserEvents2");
  Win32::OLE->Option(Warn => 4);
  $HWND_IE = $IE->{HWND};

  # Let's size the window
  $IE->{height} = $arg{height} || 600;
  $IE->{width} = $arg{width} || 808;
  $IE->{visible} = 1;

  # Show blank page (let the browser create rendering area)
  Navigate('about:blank');

  # We need the window on top because we want to get the screen shots
  Minimize($HWND_IE); Restore($HWND_IE); # this seem to work
  BringWindowToTop( $HWND_IE );

  # Find the largest child window, suppose that this is the area where the page is rendered
  my ($sz, $i) = (0, 0);
  for ( ListChilds($HWND_IE) ) {
    next unless $_->{visible};
    if ( $sz < (($_->{rect}[2]-$_->{rect}[0])*($_->{rect}[3]-$_->{rect}[1])) ) {
      $sz = (($_->{rect}[2]-$_->{rect}[0])*($_->{rect}[3]-$_->{rect}[1]));
      $i = $_->{hwnd};
    }
  }
  $HWND_Browser = $i;
}

sub QuitIE () {
  $IE->Quit();
  $IE = undef;
}

sub Navigate ($) {
  $IE->Navigate($_[0]);
  Win32::OLE->MessageLoop();
  GetDoc();
}

sub Refresh () {
  $refreshing_page = 1;
  $IE->Refresh2(3);
  Win32::OLE->MessageLoop();
  GetDoc();
}

sub GetDoc () {
  $Doc = $IE->{Document};
  $Body = $Doc->{Body};
}

sub GetElement ($) {
  return $Doc->getElementById($_[0]);
}

sub CaptureElement {
  my $e = ref $_[0] ? shift : GetElement(shift);
  return CapturePage() if $e->tagName eq 'BODY';

  my ($px, $py, $sx, $sy, $w, $h);

  # Scrolls the object so that top of the object is visible at the top of the window.
  $e->scrollIntoView();

  # This is the size of the object including border
  $w = $e->offsetWidth;
  $h = $e->offsetHeight;

  # Let's calculate the absolute position of the object on the page
  my $p = $e;
  while ( $p ) {
    $px += $p->offsetLeft;
    $py += $p->offsetTop;
    $p = $p->offsetParent;
  }

  # The position on the screen is different due to page scrolling and Body border
  $sx = $px - $Body->scrollLeft + $Body->clientLeft;
  $sy = $py - $Body->scrollTop + $Body->clientTop;

  if ( $sx+$w < $Body->clientWidth && $sy+$h < $Body->clientHeight ) {

    # If the whole object is visible
    return CaptureWindowRect($HWND_Browser,
      $sx-$CaptureBorder,
      $sy-$CaptureBorder,
      $w+2*$CaptureBorder,
      $h+2*$CaptureBorder
    );

  } else {

    # If only part of it is visible
    my (@parts, $pw, $ph, $ch, $cw);

    # We will do the screen capturing in more steps by areas of dimensions $cw x $ch
    $cw = int($Body->clientWidth * 0.8);
    $ch = int($Body->clientHeight * 0.8);

    for ( my $cnt_x=0 ; $cnt_x < ceil($w/$cw) ; $cnt_x++ ) {

      $parts[$cnt_x] = '';
      $e->scrollIntoView(); # go to object starting point
      $Doc->{parentWindow}->scrollBy($cw*$cnt_x, 0); # go to starting point for this strip

      for ( my $cnt_y=0 ; $cnt_y < ceil($h/$ch) ; $cnt_y++ ) {

        # Recalculate the position on the screen
        $sx = $px - $Body->scrollLeft + $Body->clientLeft + $cw*$cnt_x;
        $sy = $py - $Body->scrollTop + $Body->clientTop + $ch*$cnt_y;

        # Calculate the dimensions of the part to be captured
        $pw = $cw*($cnt_x+1) > $w ? $w - $cw*$cnt_x : $cw;
        $ph = $ch*($cnt_y+1) > $h ? $h - $ch*$cnt_y : $ch;

        if ( $cnt_x == 0 ) { $pw += $CaptureBorder; $sx -= $CaptureBorder; }
        if ( $cnt_y == 0 ) { $ph += $CaptureBorder; $sy -= $CaptureBorder; }
        if ( $cnt_x == floor($w/$cw) ) { $pw += $CaptureBorder; }
        if ( $cnt_y == floor($h/$ch) ) { $ph += $CaptureBorder; }

        # Capture the part and append it to the strip
        $parts[$cnt_x] .= (CaptureHwndRect($HWND_Browser, $sx, $sy, $pw, $ph))[2];

        $Doc->{parentWindow}->scrollBy(0, $ch);
      }
    }

    # join the strips into one big bitmap
    my $bw = $cw + $CaptureBorder; # width of the big bitmap
    for ( my $cnt_x=1 ; $cnt_x < ceil($w/$cw) ; $cnt_x++ ) {
      $pw = $cw*($cnt_x+1) > $w ? $w - $cw*$cnt_x + $CaptureBorder : $cw; # width of the part
      $parts[0] = JoinRawData( $bw, $pw, $h, $parts[0], $parts[$cnt_x] );
      $bw += $pw;
    }

    return CreateImage( $w+2*$CaptureBorder, $h+2*$CaptureBorder, $parts[0] );
  }
}


sub CapturePage {
  my ($px, $py, $sx, $sy, $w, $h);

  # Scrolls the object so that top of the object is visible at the top of the window.
  $Doc->{parentWindow}->scrollTo(0, 0);

  # This is the size of the page content
  $w = $Body->scrollWidth;
  $h = $Body->scrollHeight;

  # Postion is [0,0]
  $px = 0; $py = 0;

  # The position on the screen is different due to page scrolling and Body border
  $sx = $px - $Body->scrollLeft + $Body->clientLeft;
  $sy = $py - $Body->scrollTop + $Body->clientTop;

  if ( $sx+$w < $Body->clientWidth && $sy+$h < $Body->clientHeight ) {

    # If the whole object is visible
    return CaptureWindowRect($HWND_Browser, $sx, $sy, $w, $h );

  } else {

    # If only part of it is visible
    my (@parts, $pw, $ph, $ch, $cw);

    # We will do the screen capturing in more steps by areas of dimensions $cw x $ch
    $cw = int($Body->clientWidth * 0.8);
    $ch = int($Body->clientHeight * 0.8);

    for ( my $cnt_x=0 ; $cnt_x < ceil($w/$cw) ; $cnt_x++ ) {
      $parts[$cnt_x] = '';
      for ( my $cnt_y=0 ; $cnt_y < ceil($h/$ch) ; $cnt_y++ ) {

        $Doc->{parentWindow}->scrollTo($cw*$cnt_x, $ch*$cnt_y);

        # Recalculate the position on the screen
        $sx = $px - $Body->scrollLeft + $Body->clientLeft + $cw*$cnt_x;
        $sy = $py - $Body->scrollTop + $Body->clientTop + $ch*$cnt_y;

        # Calculate the dimensions of the part to be captured
        $pw = $cw*($cnt_x+1) > $w ? $w - $cw*$cnt_x : $cw;
        $ph = $ch*($cnt_y+1) > $h ? $h - $ch*$cnt_y : $ch;

        # Capture the part and append it to the strip
        $parts[$cnt_x] .= (CaptureHwndRect($HWND_Browser, $sx, $sy, $pw, $ph))[2];
      }
    }

    # join the strips into one big bitmap
    my $bw = $cw; # width of the big bitmap
    for ( my $cnt_x=1 ; $cnt_x < ceil($w/$cw) ; $cnt_x++ ) {
      $pw = $cw*($cnt_x+1) > $w ? $w - $cw*$cnt_x : $cw; # width of the part
      $parts[0] = JoinRawData( $bw, $pw, $h, $parts[0], $parts[$cnt_x] );
      $bw += $pw;
    }

    return CreateImage( $w, $h, $parts[0] );
  }
}

sub CaptureBrowser {
  $Doc->{parentWindow}->scrollTo(0, 0);
  return CaptureWindow( $HWND_IE );
}

##########################################################################

sub EventHandler {
  my ($obj,$event,@args) = @_;

  if ($event eq 'DocumentComplete' && $IE->ReadyState() == 4)  {
    Win32::OLE->QuitMessageLoop;
  }

  if ($event eq 'DownloadComplete' && $refreshing_page) {
    $refreshing_page = 0;
    Win32::OLE->QuitMessageLoop;
  }
}

##########################################################################

1;

__END__

=head1 NAME

Win32::CaptureIE - Capture web pages or its elements rendered by Internet Explorer

=head1 SYNOPSIS

  use Win32::CaptureIE;

  StartIE;
  Navigate('http://my.server/page.html');

  my $img = CaptureElement('tab_user_options');
  $img->Write("ie-elem.png");

  QuitIE;

=head1 DESCRIPTION

The package enables you to automatically create screenshots of your
web server pages for the user guide or whatever you need it for. The
best part is that you don't bother yourself with scrolling and object
localization. Just tell the ID of the element and receive an Image::Magick
object. The package will do all the scrolling work, it will take the
screenshots and glue the parts together.

=head1 EXPORT

=over 8

=item :default

C<CaptureBrowser>
C<CaptureElement>
C<CapturePage>
C<GetDoc>
C<GetElement>
C<Navigate>
C<QuitIE>
C<Refresh>
C<StartIE>
C<$Body>
C<$CaptureBorder>
C<$Doc>
C<$HWND_Browser>
C<$HWND_IE>
C<$IE>

=back

=head2 Internet Explorer controlling functions

=over 8

=item StartIE ( %params )

This function creates a new Internet Explorer process via Win32::OLE.
You can specify width and height of the window as parameters.

  StartIE( width => 808, height => 600 );

The function will bring the window to the top and try to locate the
child window where the page is rendered.

=item QuitIE ( )

Terminates the Internet Explorer process and destroys the Win32::OLE object.

=item Navigate ( $url )

Loads the specified page and waits until the page is completely loaded. Then it will
call C<GetDoc> function.

=item Refresh ( )

Refreshes the currently loaded page and calls C<GetDoc> function.

=item GetDoc ( )

Loads C<$Doc> and C<$Body> global variables.

=item GetElement ( $id )

Returns the object of specified ID by calling C<< document->getElementById() >>.

=back

=head2 Capturing functions

These function works like other C<Capture*(...)> functions from L<Win32::Screenshot|Win32::Screenshot> package.

=over 8

=item CaptureBrowser ( )

Captures whole Internet Explorer window including the window title and border.

=item CapturePage ( )

Captures whole page currently loaded in the Internet Explorer window. Only the page content will
be captured - no window, no scrollbars. If the page is smaller than the window only the occupied
part of the window will be captured. If the page is longer (scrollbars are active) the function
will capture the whole page step by step by scrolling the window content (in all directions) and
will return an complete image of the page.

=item CaptureElement ( $id | $element )

Captures the element specified by its ID or passed as reference to the
element object. It will capture a small border around the element
specified by C<$CaptureBorder> global variable. The function will
scroll the page content to show the top of the element and scroll down
and right step by step to get whole area occupied by the object.

=back

=head2 Global variables

=over 8

=item $CaptureBorder

The function C<CaptureElement> is able to capture the element and
a small area around it. How much of surrounding space will be captured
is defined by C<$CaptureBorder>. It is not recommended to capture more
than 3-5 pixels because parts of other elements could be captured
as well. Default border is 1 pixel wide.

=item $IE

The function C<StartIE> will create a new Internet Explorer process
and its Win32::OLE reference will be stored in this variable. See the
MSDN documentation for InternetExplorer object.

=item $Doc

The function C<GetDoc> will assign C<< $IE->{Document} >> into this
variable. See the MSDN documentation for Document object.

=item $Body

The function C<GetDoc> will assign C<< $IE->{Document}->{Body} >> into this
variable. See the MSDN documentation for BODY object.

=item $HWND_IE

The function C<StartIE> will assign the handle of the Internet Explorer window
into this variable from C<< $IE->{HWND} >>.

=item $HWND_Browser

The function C<StartIE> will try to find the largest child window and
suppose that this is the area where is the page rendered. It is used to
convert page coordinates to screen coordinates.

=back

=head1 BUGS

=over 8

=item * Access denied

Sometimes I receive the error message 'Access denied' when accessing
C<< $Doc->{parentWindow} >> properties and methods (C<scrollBy> and
C<scrollTo> methods are called from the package functions). Sometimes
I have to close all running IE windows or restart computer if it's not
enough to solve the problem. I will appreciate any help.

=back

=head1 SEE ALSO

=item MSDN

http://msdn.microsoft.com/library You can find there the description
of InternetExplorer object and DOM.

=item L<Win32::Screenshot|Win32::Screenshot>

This package is used for capturing screenshots. Use its post-processing
features for automatic screenshot modification.

=head1 AUTHOR

P.Smejkal, E<lt>petr.smejkal@seznam.czE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2004 by P.Smejkal

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.8.2 or,
at your option, any later version of Perl 5 you may have available.


=cut

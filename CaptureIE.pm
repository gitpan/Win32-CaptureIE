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
  CaptureRows
  CaptureThumbshot

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

our $VERSION = '1.11';
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

# HACK: DocumentComplete event is not fired when refreshing page but
# only DownloadComplete event, so if refreshing page we need to wait for
# DownloadComplete but not if we are navigating to a page
our $refreshing_page = 0;

sub StartIE {
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
  $Body = (! $Doc->compatMode || $Doc->compatMode eq 'BackCompat') ? $Doc->{Body} : $Doc->{Body}->{parentNode};
}

sub GetElement ($) {
  return $Doc->getElementById($_[0]);
}


sub CaptureRows {
  my $tab = ref $_[0] ? shift : GetElement(shift);
  my %rows = map {$_ => 1} ref $_[0] ? @{$_[0]} : @_;
  return undef if $tab->tagName ne 'TABLE' || !%rows;

  my $img;
  {
    local @Win32::Screenshot::POST_PROCESS = ();
    $img = CaptureElement($tab);
  }

  my $pos = $CaptureBorder + $tab->rows(0)->{offsetTop};

  for ( my $row = 0 ; $row < $tab->rows->{length} ; $row++ ) {
    if ( $rows{$row} ) {
      $pos += $tab->rows($row)->{offsetHeight};
    } else {
      $img->Chop('x'=>0, 'y'=>$pos, 'width'=>0, 'height'=>$tab->rows($row)->{offsetHeight});
    }
  }

  return PostProcessImage( $img );
}


sub CaptureThumbshot {

  GetDoc();

  # resize the window to set the client area to 800x600
  $IE->{width} = $IE->{width} + 800-$Body->clientWidth;
  $IE->{height} = $IE->{height} + 600-$Body->clientHeight;

  # scrollTo(0, 0)
  $Body->doScroll('pageUp') while $Body->scrollTop > 0;
  $Body->doScroll('pageLeft') while $Body->scrollLeft > 0;

  Win32::OLE->SpinMessageLoop();

  return CaptureWindowRect($HWND_Browser, $Body->clientLeft, $Body->clientTop, $Body->clientWidth, $Body->clientHeight );
}


sub CaptureElement {
  my $e = ref $_[0] ? shift : GetElement(shift);
  my %args = ref $_[0] eq 'HASH' ? %{(shift)} : ();
  return CapturePage() if $e->tagName eq 'BODY';

  GetDoc();

  $args{border_left} = exists $args{border_left} ? exists $args{border_left} : exists $args{border} ? $args{border} : $CaptureBorder;
  $args{border_right} = exists $args{border_right} ? exists $args{border_right} : exists $args{border} ? $args{border} : $CaptureBorder;
  $args{border_top} = exists $args{border_top} ? exists $args{border_top} : exists $args{border} ? $args{border} : $CaptureBorder;
  $args{border_bottom} = exists $args{border_bottom} ? exists $args{border_bottom} : exists $args{border} ? $args{border} : $CaptureBorder;

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

  $px -= ($args{border_left}||0);
  $py -= ($args{border_top}||0);
  $w  += ($args{border_left}||0) + ($args{border_right}||0);
  $h  += ($args{border_top}||0) + ($args{border_bottom}||0);

  # The position on the screen is different due to page scrolling and Body border
  $sx = $px - $Body->scrollLeft + $Body->clientLeft;
  $sy = $py - $Body->scrollTop + $Body->clientTop;

  if ( $sx+$w < $Body->clientWidth && $sy+$h < $Body->clientHeight ) {

    # If the whole object is visible
    return CaptureWindowRect($HWND_Browser, $sx, $sy, $w, $h );

  } else {

    # If only part of it is visible
    return CaptureAndScroll($e, $px, $py, $w, $h);
  }
}

sub CapturePage {
  my ($px, $py, $sx, $sy, $w, $h);

  GetDoc();

  # Scrolls the object so that top of the object is visible at the top of the window.
  $Body->doScroll('pageUp') while $Body->scrollTop > 0;
  $Body->doScroll('pageLeft') while $Body->scrollLeft > 0;

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
    return CaptureAndScroll(undef, $px, $py, $w, $h);
  }
}

sub CaptureAndScroll {
  my ($e, $px, $py, $w, $h) = @_;
  my ($strip, $final, $pw, $ph, $ch, $cw, $maxw, $maxh, $sx, $sy);

  GetDoc();

  $final = '';

  # Captured area
  $cw = 0;
  $ch = 0;

  # We will do the screen capturing in more steps by areas of maximum dimensions $cw x $ch
  $maxw = $Body->clientWidth;
  $maxh = $Body->clientHeight;

  for ( my $cnt_x=0 ; $cw < $w ; $cnt_x++ ) {

    # Scroll to the top and one right
    if ( $e ) {
      $e->scrollIntoView;
      $Body->doScroll('pageRight') for 1..$cnt_x;
    } else {
      $Body->doScroll('pageUp') while $Body->scrollTop > 0;
      $Body->doScroll('pageRight') if $cnt_x;
    }
    Win32::OLE->SpinMessageLoop;

    $strip = '';
    $ch = 0;

    for ( my $cnt_y=0 ; $ch < $h ; $cnt_y++ ) {

      $Body->doScroll('pageDown') if $cnt_y;

      # Recalculate the position on the screen
      $sx = $px - $Body->scrollLeft + $Body->clientLeft + $cw;
      $sy = $py - $Body->scrollTop + $Body->clientTop + $ch;

      # Calculate the dimensions of the part to be captured
      $pw = ($px+$cw) - $Body->scrollLeft + $maxw > $maxw ? $maxw - ($px+$cw) + $Body->scrollLeft : $maxw;
      $pw = $cw + $pw > $w ? $w - $cw : $pw;

      $ph = ($py+$ch) - $Body->scrollTop + $maxh > $maxh ? $maxh - ($py+$ch) + $Body->scrollTop : $maxh;
      $ph = $ch + $ph > $h ? $h - $ch : $ph;

      # Capture the part and append it to the strip
      $strip .= (CaptureHwndRect($HWND_Browser, $sx, $sy, $pw, $ph))[2];

      $ch += $ph;
    }

    $final = JoinRawData( $cw, $pw, $h, $final, $strip );

    $cw += $pw;
  }

  return CreateImage( $w, $h, $final );
}


sub CaptureBrowser {

  GetDoc();

  # scrollTo(0, 0)
  $Body->doScroll('pageUp') while $Body->scrollTop > 0;
  $Body->doScroll('pageLeft') while $Body->scrollLeft > 0;

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
C<CaptureRows>
C<CapturePage>
C<CaptureThumbshot>
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
will return a complete image of the page.

=item CaptureElement ( $id | $element [, \%args ] )

Captures the element specified by its ID or passed as reference to the
element object. The function will scroll the page content to show the top
of the element and scroll down and right step by step to get whole area
occupied by the object.

It can capture a small surrounding area around the element specified
by %args hash or C<$CaptureBorder> global variable. It recognizes paramters
C<border>, C<border-left>, C<border-top>, C<border-right> and C<border-bottom>.
The priority is C<border-*> -> C<border> -> C<$CaptureBorder>.

=item CaptureRows ( $id | $element , @rows )

Captures the table specified by its ID or passed as reference to the
table object. The function will scroll the page content to show the top
of the table and scroll down and right step by step to get whole area
occupied by the table. Than it will chop unwanted rows from the image and it will
return the image of table containing only selected rows. Rows are numbered from zero.

It can capture a small surrounding area around the element specified
by C<$CaptureBorder> global variable.

=item CaptureThumbshot ( )

Resizes the window to set the client area to 800x600 pixels. Captures the client
area where the page is rendered. No scrolling is done.

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

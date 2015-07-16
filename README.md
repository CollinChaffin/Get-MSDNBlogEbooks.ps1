Get-MSDNBlogEbooks.ps1
============
```
#==================================================================================================#
#
# Author:				Collin Chaffin
# Last Modified:		07-12-2015 01:00 PM
# Filename:				Get-MSDNBlogEbooks.ps1
#
#
# Changelog:
#
#	v 1.0.0.1	:	07-16-2015	:	Initial release
#
# Notes:
#
#	This script automates downloading all free ebooks posted to the
#	following Microsoft MSDN blog:  http://bit.ly/FreeMSDNbooks
#
#	It Optionally will create a local zip file for archiving containing
#	all the downloaded ebook files.
#
#	Example:
#		Get-MSDNBlogEbooks.ps1 -SavePath "C:\Data\eBooks" -CreateZip -ZipFile "C:\Temp\eBooks.zip"
#
#
#
#    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
#    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
#    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
#    PARTICULAR PURPOSE AND NONINFRINGEMENT.
#
#    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
#    FOR ANY SPECIAL, INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY
#    DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
#    WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS
#    ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE
#    OF THIS CODE OR INFORMATION.
#
#
#
#	**NOTE -- NO EDITING SHOULD BE REQUIRED IF THE REQUIRED PARAMS ARE PROVIDED ON COMMAND LINE**
#	**NOTE -- AT TIME OF AUTHORING, SEVERAL OF THE PROVIDED EBOOK URLS WERE NOT REACHABLE**
#	**NOTE -- THIS DOWNLOADS OVER 1.2 GIGABYTES OF EBOOKS.  EVEN ZIPPING IS OVER A GIG!!!!!!!!
#
#==================================================================================================#
```
# Direct Membership Collection

## Overview
Sample/prototype of creating a direct membership collection in SCCM.  Reads a CSV file with the host names (NETBIOS name/short name) for the client and
create a collection containing each host.

The configuration manager console is required to run this code.  You must also have valid credentials to access your SCCM Site and create/modify collections.

## Usage (Command line arguments)

Either build and execute directly or configure command line arguments in Visual Studio and Build/run.

Argument list as follows/example:

```
DirectCollectionMembership.exe fqdn.of.your.site.server userName F:\ull\Path\to.csv intForHostNameFieldInInputFile(0) LimitCollectionId "Name of new Collection"
```
You'll be prompted to input your password.

See source code in Program.cs for details.


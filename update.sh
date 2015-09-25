# Script for update content of server folder in mac
SERVER="/Library/WebServer/Documents/"
rm -rf $SERVER/*
cp -r * $SERVER
chmod -R 777 $SERVER/register_xls

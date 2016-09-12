#include<fstream>
#include<iostream>
using namespace std;
ifstream inf ("CS.txt");
ofstream ouf ("CO.txt");
int main()
{
    string m[20902+121]; 
    int i;
    for(i=0;i<=20901+121;i++)
     inf>>m[i];
    for(i=0;i<=20901+121;i++)
      ouf<<"\""<<m[i]<<"\""<<endl;
    cout<<"Change Over!";
    system("pause");
    return 0;
}

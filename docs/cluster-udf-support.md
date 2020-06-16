---
layout: post
title: "Offloading UDF computations to a Windows HPC cluster from Excel 2010"
date: 2010-12-09 21:02:00 -0000
permalink: /2010/12/09/cluster-udf-support/
categories: features, .net, excel, exceldna, hpc, xll
---
Excel 2010 introduced support for offloading UDF computations to a compute cluster. The Excel blog talks about it [https://www.microsoft.com/en-us/microsoft-365/blog/2010/02/12/offloading-udfs-to-a-windows-hpc-cluster/][hpc-cluster], and there are some nice pictures on this TechNet article: [http://technet.microsoft.com/en-us/library/ff877825(WS.10).aspx][hpc-services].

Excel-DNA now supports marking functions as cluster-safe, and I have updated the loader to allow add-ins to work under the `XllContainer` on the HPC nodes. There are some issues to be aware of:

* The add-in does not create its own `AppDomain` when running on the compute node. One consequence is that no custom `.xll.config` file is used; configuration entries need to be set in the `XllContainer` configuration setup.
* There are some limitations on the size of array data that can be passed to and from UDF calls - this limit is probably configurable in the WCF service.
* Only the 32-bit host is currently supported.

To test this you will need an Windows HPC Server 2008 R2 cluster with the HPC Services for Excel installed. On the clients you need Excel 2010 with the HPC cluster connector installed. The latest check-in for Excel-DNA with this support is on GitHub: [https://github.com/Excel-DNA/ExcelDna][main-repo].

In the Microsoft HPC SDK there is a sample called ClusterUDF.xll with a few test functions. I have recreated these in C# in the samples file [Distribution\Samples\ClusterSample.dna][cluster-sample] Basically functions just need to be marked as `IsClusterSafe=true` to be pushed to the cluster for computation. For example

{% highlight csharp %}
[ExcelFunction(IsClusterSafe=true)]
public static int DnaCountPrimesC(int nFrom, int nTo)
{
    // ...
}
{% endhighlight %}

As usual, any feedback on this feature - questions or reports on whether you use it - will be most appreciated.

[hpc-cluster]: https://www.microsoft.com/en-us/microsoft-365/blog/2010/02/12/offloading-udfs-to-a-windows-hpc-cluster/
[hpc-services]: http://technet.microsoft.com/en-us/library/ff877825(WS.10).aspx
[main-repo]: https://github.com/Excel-DNA/ExcelDna
[cluster-sample]: https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/Samples/ClusterSample.dna

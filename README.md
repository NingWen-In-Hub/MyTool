# MyTool
工作的时候随手制作的小工具

HtmlCleaner.py……针对radio button写得令人强迫症发作的html的修正工具。
  对input的属性进行排序，# type > style > class > id > name > else
  检查不同组(同一个td下)的radio button的name属性是否重复，并在重复时发出警报
  对同组(同一个td下)的radio button进行id的覆写，根据name属性分配id,并把新id同步到邻接的label.for属性上
  最后，享受你的新Html

vocabularySizeAnalysis.py……根据给出文章估算读懂需要的词汇量。配合dic.csv食用

# Yelp Search

Script that accepts a list of businesses and outputs an Excel file of Yelp search 

## Usage

`python yelp_search.py <filename> <zipcode>`

Example:

`python yelp_search.py RamenList.txt 90001`

RamenList.txt:

```
Santouka
...
```

Console output:

```
Santouka
Hokkaido Ramen Santouka
4.5
Number of reviews: 1971
3760 S Centinela Ave
Los Angeles, CA 90066
+1-310-391-1101
https://www.yelp.com/biz/hokkaido-ramen-santouka-los-angeles?adjust_creative=mUML_rr9Gv6-dmefggsCHg&utm_campaign=yelp_api&utm_medium=api_v2_search&utm_source=mUML_rr9Gv6-dmefggsCHg

...
```

Excel file:

Search | Name	| Rating	| Number of Reviews	| Street Address	| City	| State |	Zip Code | Phone Number	| Yelp Link
--- | --- | --- | --- | --- | --- | --- | --- | --- | ---
Santouka | Hokkaido Ramen Santouka | 4.5 | 1971 | 1971	3760 S Centinela Ave| Los Angeles | CA	| 90066	| +1-310-391-1101	| https://www.yelp.com/biz/hokkaido-ramen-santouka-los-angeles?adjust_creative=mUML_rr9Gv6-dmefggsCHg&utm_campaign=yelp_api&utm_medium=api_v2_search&utm_source=mUML_rr9Gv6-dmefggsCHg
... | ... | ... | ... | ... | ... | ... | ... | ... | ...

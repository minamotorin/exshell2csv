#!/usr/bin/env awk -f

# field n and after n
function after(n,                                                   r) {
  r="";
  for(m=n;m<=NF;m++) r = r " " $m;
  sub(" ", "", r);
  return r;
}

# step ALPHABET counter (e.g. ABZ -> ACA)
function ALPH_advance(s,                                         a, p) {
  a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_";
  p = length(s)
  while (p != -1) {
    match(a, substr(s, p, 1));
    s = substr(s, 1, p-1) substr(a, RSTART+1, 1) substr(s, p+1, length(s));
    match(s, "_")
    p = RSTART - 1;
    sub("_", "A", s);
    if (p==0) {
      s = "A" s;
      break;
    }
  }
  return s;
}

# n < m or not (n, m: ALPHABET counter)
function ALPH_lt(n, m) {
  if (n=="") return 1;
  if (length(n) > length(m)) return 0;
  if (length(n) < length(m)) return 1;
  for(l=1;n;l++) {
    if(substr(n, 1, 1) == substr(m, 1, 1)) {
      sub(".", "", n);
      sub(".", "", m);
    } else if (substr(n, 1, 1) < substr(m, 1, 1)) return 1; # success due to a letter
      else return 0;
  }
  return 0; # n == m
}

# format serial number to Y/M/D/A
function fdate(n,                                          y, m, d, a) {
  # n: count of number ou dates since 1900
  m = n;
  n = n + 1900*365+475-19+5; # until 1900
  n = n - 2; # 1900 is not a leap year (differ from Microsoft Excel)
  y = sprintf("%d", n/365);
  n = n - y * 365;
  n = n - sprintf("%d", y / 4);
  n = n + sprintf("%d", y / 100);
  n = n - sprintf("%d", y / 400);
  if (y%4==0&&(y%400==0||y%100!=0)) a = 1;
  while (n <= 0) {
    n = n + 365 + a;
    y = y -1;
    if (y%4==0&&(y%400==0||y%100!=0)) a = 1;
  }

  for(M=1;M<12;M++) {
    if(M==2) d = 28 + a;
    else if (M<8) {
      if (M%2==0) d=30;
      else d = 31;
    } else {
      if (M%2==0) d = 31;
      else d = 30;
    }
    if (n > d) n = n - d;
    else break;
  }

  a = m;
  m = M;
  d = n;
  a = (a + 6) % 7;# sun is 0
  return y"/"m"/"d"/"a;
}

# from sharedStrings.xml
$1=="l"{
  str[NR-1] = after(2);
  gsub("\\\\n", "\n", str[NR-1]);
  gsub("\\\\\\\\", "\\\\", str[NR-1]);
}

# from sheet
$1!="l"{
  if (ALPH_lt(ymax, $1)) ymax = $1;
  if (xamx < $2) xmax = $2;

  if ($3=="s") {# refer to sharedStrings.xml
    cell[$2, $1] = str[$4];
  }
############################ CUSTOMIZE AREA ############################
  else if ($3==45||$3==40) {# time H:M
    t=sprintf("%d", $4*24);
    t=($4*24-t)*60;
    t=sprintf("%d", t);
    if (length(t)==1) t = "0"t;
    cell[$2, $1] = sprintf("%d", $4*24)":"t;
  }
  else if ($3==16) {# date Y/M/D
    cell[$2, $1] = fdate($4);
    sub("/.$", "", cell[$2, $1]);
  }
########################## END CUSTOMIZE AREA ##########################
  else {# others
    cell[$2, $1] = after(4);
  }
}

#output CSV
END{
  debug = 0; # add column and line numbers
  if (debug) {
    printf("%s,", 0);
    for (n="A";ALPH_lt(n, ymax);n=ALPH_advance(n)) printf("%s,", n);
    print ymax;
  }

  for (m=1;m<=xmax;m++) {
    if (debug) printf("%s,", m);
    for (n="A";ALPH_lt(n, ymax);n=ALPH_advance(n)) {
      c =  cell[m, n];
      gsub("\"", "\"\"", c);
      if (match(c, ",")||match(c, "\\n")||match(c, "\"")) printf("\"%s\",", c);
      else printf("%s,", c);
    }
      c =  cell[m, ymax];
      gsub("\"", "\"\"", c);
      if (match(c, ",")||match(c, "\\n")||match(c, "\"")) print "\""c"\"";
      else print c;
  }
}

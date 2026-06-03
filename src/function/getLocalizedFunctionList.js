import functionlist from "./functionListDescriptor";

export default localeFunctionList => {
  functionlist.forEach(f => {
    const localeFunction = localeFunctionList[f.n];
    if (localeFunction) {
      f.d = localeFunction.d;
      f.a = localeFunction.a;
      if (localeFunction.p) {
        f.p.forEach((p, i) => {
          if (localeFunction.p[i]) {
            Object.assign(p, localeFunction.p[i]);
          }
        });
      }
    }
  });

  return functionlist;
};

import struct, sys
class CFB:
    def __init__(self, path):
        self.d=open(path,'rb').read(); d=self.d
        assert d[:8]==bytes.fromhex('d0cf11e0a1b11ae1')
        self.ssz=1<<struct.unpack_from('<H',d,30)[0]
        self.mssz=1<<struct.unpack_from('<H',d,32)[0]
        nfat=struct.unpack_from('<I',d,44)[0]
        self.dirstart=struct.unpack_from('<I',d,48)[0]
        self.cutoff=struct.unpack_from('<I',d,56)[0]
        self.minifatstart,self.nminifat,difstart,ndif=struct.unpack_from('<IIII',d,60)
        difat=list(struct.unpack_from('<109I',d,76))
        s=difstart
        for _ in range(ndif):
            sec=self.sec(s); difat+=list(struct.unpack_from('<%dI'%(self.ssz//4-1),sec,0)); s=struct.unpack_from('<I',sec,self.ssz-4)[0]
        fat=[]
        for fs in difat[:nfat]: fat+=list(struct.unpack_from('<%dI'%(self.ssz//4),self.sec(fs)))
        self.fat=fat
        self.entries=self._dirs()
        root=[e for e in self.entries if e[2]==5][0]
        self.ministream=self._read_chain(root[3])
        mf=self._read_chain(self.minifatstart)
        self.minifat=list(struct.unpack_from('<%dI'%(len(mf)//4),mf)) if mf else []
    def sec(self,i): return self.d[(i+1)*self.ssz:(i+2)*self.ssz]
    def chain(self,start,fat=None):
        fat=fat or self.fat; out=[];s=start
        while s<0xFFFFFFFA and s<len(fat): out.append(s); s=fat[s]
        return out
    def _read_chain(self,start):
        return b''.join(self.sec(s) for s in self.chain(start))
    def _dirs(self):
        data=self._read_chain(self.dirstart); out=[]
        for i in range(len(data)//128):
            e=data[i*128:(i+1)*128]; nlen=struct.unpack_from('<H',e,64)[0]
            name=e[:max(0,nlen-2)].decode('utf-16le','replace')
            out.append((i,name,e[66],struct.unpack_from('<I',e,116)[0],struct.unpack_from('<Q',e,120)[0]))
        return out
    def read(self,name):
        for i,n,t,st,sz in self.entries:
            if n==name and t==2:
                if sz<self.cutoff:
                    out=b'';s=st
                    while s<0xFFFFFFFA and s<len(self.minifat):
                        out+=self.ministream[s*self.mssz:(s+1)*self.mssz]; s=self.minifat[s]
                    return out[:sz]
                return self._read_chain(st)[:sz]
        raise KeyError(name)
for p in sys.argv[1:]:
    c=CFB(p); print("==",p)
    for i,n,t,st,sz in c.entries:
        if t: print(f"  [{i}] {n!r} type={t} start={st} size={sz}")

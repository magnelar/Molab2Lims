--12:19 08.03.2017
SELECT   i.instr_code,p.meth_org_unit, p.methcode, m.meth_name, p.parord, p.parno, p.unit,
         p.NAME, p.DEC, p.rep_level, f.from_meth_code, f.seq_no,
         F.FROM_PAR_NO, F.FROM_UNIT, F.FROM_NAME, F.CALC_TYPE, F.FACTOR
    FROM elp_meth m, elp_parmet p, elp_formel f,elp_instr i
   WHERE M.METH_ORG_UNIT = 'BRE'
     AND M.METHCODE = 'SILICAZMAN'
     and m.instr_code=i.instr_code(+)
     AND m.meth_org_unit = p.meth_org_unit
     AND m.methcode = p.methcode
     AND p.meth_org_unit = f.meth_org_unit(+)
     AND p.methcode = f.meth_code(+)
     AND p.parno = f.par_no(+)
     AND NVL (f.delete_flag, 'N') = 'N'
     AND NVL (p.delete_flag, 'N') = 'N'
     AND REG_METHOD='K'
ORDER BY m.methcode, parord, f.seq_no

EXECUTE CALCULATE.ANALYSE_CALCULATE('BRE',201708646);

select f.seq_no,f.from_meth_org_unit,
          f.from_meth_code,f.from_par_no,f.calc_type,
          f.calc_ref,nvl(p.reg_method,'F'),f.factor
          from elp_formel f,elp_parmet p
          where p.meth_org_unit (+)=f.from_meth_org_unit
            and p.methcode (+)=f.from_meth_code
            AND P.PARNO (+)=F.FROM_PAR_NO
            and f.meth_org_unit= 'BRE'
            AND F.METH_CODE ='SILICAZMAN'
            and f.par_no = 114
            AND NVL(F.DELETE_FLAG,'N') = 'N'
          order by f.seq_no;
		  
select r.elms_org_unit,r.elmsno, r.parno,r.meth_org_unit,r.methcode,r.value, r.rowid
            FROM ELD_RES R
            where r.elms_org_unit='BRE'
              AND R.ELMSNO=201708646
              AND R.INPUT_METHODE ='K'
              and nvl(r.delete_flag,'N') = 'N';
			  
SELECT m.meth_org_unit,i.instr_org_unit,
 m.instr_code,m.methcode ,  p.ident_code, samp_column, from_pos, to_pos,to_pos-from_pos+1 lengde, format_string
    FROM ELP_IDENT P,ELP_METH M,ELP_INSTR I
   WHERE M.INSTR_CODE = I.INSTR_CODE  AND I.IDENT_CODE = P.IDENT_CODE 
      --AND m.methcode = 'SILICA'
	  
SELECT to_pos-from_pos+1 lengde  FROM ELP_IDENT P,ELP_METH M,ELP_INSTR I  WHERE M.INSTR_CODE = I.INSTR_CODE  AND I.IDENT_CODE = P.IDENT_CODE AND m.methcode = 'SILICA'
--09:55 10.03.2017
-- se C:\Users\ismel\Dropbox\GD_Jobb\YoungDeng\DiverseSql.sql
         OVN1  TAPP20170001      1A 280217  1701131201      
SILICA-Z
		lSMess	"         OVN1  TAPP20170001  1A   2802171701131201"	String
		lSMess	"         OVN1  TAPP20170001  1A   2802171701131201 "	String

SELECT   m.methcode,p.ident_code, samp_column, from_pos, to_pos, format_string,to_pos-from_pos+1 lengde,
         SUBSTR ('         OVN1  TAPP20170001  1A   2802171701131201',
                 from_pos,
                 TO_POS - FROM_POS + 1
                ) ident
    FROM elp_ident p, elp_meth m, elp_instr i
   WHERE m.instr_code = i.instr_code
     AND i.ident_code = p.ident_code
     AND m.methcode = 'SILICA-Z'
ORDER BY from_pos;

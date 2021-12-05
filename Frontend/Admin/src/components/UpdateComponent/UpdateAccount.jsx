/* eslint-disable jsx-a11y/alt-text */
import Checkbox from '@mui/material/Checkbox';
import Fade from '@mui/material/Grow';
import Paper from '@mui/material/Paper';
import { useSnackbar } from 'notistack';
import React, { useContext, useEffect, useState } from 'react';
import { getProductId, updateProduct } from '../../app/ApiResult';
import { context } from '../../app/Context';
import Account from './../Account/index';
import './stylesUpdateComponent/UpdateAccount.scss';
function UpdateAccount(props) {
  const Context = useContext(context);
  const { id } = props; 
const { enqueueSnackbar } = useSnackbar();
  const { setBodyAdmin, setFillerAdmin,TypeDataPro } = Context;
  const [valueData, setValueData] = useState({
    Id: '',
    Username: '',
    Password: '',
    Role: '',
  });
   
  // eslint-disable-next-line react-hooks/exhaustive-deps
  // useEffect(async() => {
  //   const result = await getProductId(id,"/product")
  //   console.log(result)
  // if(result){

  //   setValueData({
  //     ...valueData,
  //    Id:result.Id,
  //    Username:result?.Name,
  //    Password:result?.Description,
  //    Role:result?.Description,
  //   })
  // }

  // eslint-disable-next-line react-hooks/exhaustive-deps
  // },[id])

  function Prev() {
    setBodyAdmin(<Account />);
    setFillerAdmin('ACCOUNT');
  }

  const handleChange = (event) => {
    setValueData({ ...valueData, [event.target.name]: [event.target.value].toString() });

  }
  return (
    <div className='UpdateAccount'>
      <Fade in={true} timeout={200} style={{ height: '100%' }}>
        <Paper>
          <button
             style={{width:'fit-content', position:'absolute'}}
            type='button'
            className='btn btn-success d-flex gap-2'
            onClick={() => Prev()}>
            <i
              style={{ fontSize: '1.5rem' }}
              className='fad fa-chevron-circle-left'></i>
            <p className> Quay lại</p>
          </button>
          <h2 className='text-center pt-4 '>Cập nhật tài khoản </h2>

          <div className='dataAdd'>
            <div className='form-floating mb-3 inputData'>
              <input
                type='text'
                className='form-control '
                name='Username'
                color='warning'
                value={valueData?.Username}
                onChange={handleChange}
              />
              <label htmlFor='floatingInput'>Username</label>
            </div>

            <div className='form-floating mb-3 inputData'>
              <input
                type='text'
                className='form-control'
                name='Password'
                color='warning'
                value={valueData?.Password}
                onChange={handleChange}
              />

              <label htmlFor='floatingInput'>Password</label>
            </div>
            <div className='form-floating mb-3 inputData'>
              <select
                type='text'
                className='form-control '
                name='Role'
                color='warning'
                value={valueData?.Role}
                onChange={handleChange}>
                <option value='1'>1</option>
                <option value='2'>2</option>
                <option value='3'>3</option>
              </select>

              <label htmlFor='floatingInput'>Role</label>
            </div>

            <div className='inputData'>
              <button
                type='submit'
                className='btn btn-success inputData'
                style={{ width: '100%', margin: '0 auto' }}
                >
             Cập nhật tài khoản
              </button>
            </div>
          </div>
        </Paper>
      </Fade>
    </div>
  );
}

export default UpdateAccount;

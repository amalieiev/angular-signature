import { Component } from '@angular/core';
import { getAccessTokenAsync } from 'src/app/shared/office';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  public async onLogin(): Promise<void> {
    const token = await getAccessTokenAsync();
    console.log(token);
  }
}
